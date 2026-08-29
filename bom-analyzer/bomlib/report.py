"""Turn analysis results into the tables the CLI writes and prints.

One column layout feeds the CSV writer, the .xlsx writer and the terminal
table, so the three can never drift apart.
"""

import csv
import datetime
import json
import os
import re

from . import leadtime as leadtime_module
from .normalize import format_lead_time
from .prepare import DUPLICATE, FLAGGED, IGNORED, MERGED
from .xlsx_writer import (
    Cell,
    STYLE_BAD,
    STYLE_BOLD,
    STYLE_DEFAULT,
    STYLE_GOOD,
    STYLE_HEADER,
    STYLE_INT,
    STYLE_INT_BOLD,
    STYLE_LABEL,
    STYLE_MONEY,
    STYLE_MONEY_BOLD,
    STYLE_MONEY_FINE,
    STYLE_MUTED,
    STYLE_SECTION,
    STYLE_SUBTITLE,
    STYLE_TITLE,
    STYLE_WARN,
    column_letter,
    fill_style,
    write_xlsx,
)

BASE_COLUMNS = ['Row', 'Part Number', 'Quantity', 'Reference', 'Manufacturer', 'Description']
BASE_WIDTHS = [6, 26, 9, 16, 18, 34]

SUPPLIER_COLUMNS = [
    'P/N', 'Stock', 'Lead Time', 'Lead Days', 'Unit Price',
    'Extended Price', 'Order Qty', 'Lifecycle', 'Status',
]
SUPPLIER_WIDTHS = [22, 11, 12, 10, 12, 14, 10, 14, 26]

TAIL_COLUMNS = [
    'Cheapest Supplier', 'Soonest Supplier', 'Soonest (days)',
    'Recommended', 'Worst Lifecycle', 'Availability', 'Notes',
]
TAIL_WIDTHS = [17, 20, 14, 14, 18, 22, 52]

SEVERITY_STYLE = {'bad': STYLE_BAD, 'warn': STYLE_WARN, 'ok': STYLE_GOOD, 'unknown': STYLE_MUTED}


def build_header(suppliers, currency='USD'):
    header = list(BASE_COLUMNS)
    for supplier in suppliers:
        for column in SUPPLIER_COLUMNS:
            if column in ('Unit Price', 'Extended Price'):
                header.append('%s %s (%s)' % (supplier['name'], column, currency))
            else:
                header.append('%s %s' % (supplier['name'], column))
    header.extend(TAIL_COLUMNS)
    return header


def column_widths(suppliers):
    widths = list(BASE_WIDTHS)
    for _ in suppliers:
        widths.extend(SUPPLIER_WIDTHS)
    widths.extend(TAIL_WIDTHS)
    return widths


def build_rows(result, summary, styled=False):
    """Return [[value|Cell, ...], ...] — the header row plus one row per part."""
    suppliers = result['suppliers']
    currency = summary.get('currency') or 'USD'
    # The audit sheet also feeds the CSV writer, which has no colour at all, so
    # the band travels as a word — otherwise the one thing the other exports
    # say at a glance is the one thing this one loses.
    bands = leadtime_module.bands_by_row(result)

    def cell(value, style=STYLE_DEFAULT):
        return Cell(value, style) if styled else value

    rows = [[cell(name, STYLE_HEADER) for name in build_header(suppliers, currency)]]

    for index, row in enumerate(result['rows']):
        comparison = row['comparison']
        record = [
            cell(row.get('row'), STYLE_INT),
            cell(row.get('mpn')),
            cell(row.get('quantity'), STYLE_INT),
            cell(row.get('reference')),
            cell(row.get('manufacturer')),
            cell(row.get('description')),
        ]

        for supplier in suppliers:
            offer = row['offers'].get(supplier['id'])
            if not offer or not offer.get('found'):
                reason = (offer or {}).get('reason') or 'No match'
                style = STYLE_BAD if (offer or {}).get('error') else STYLE_MUTED
                record.extend([cell(None)] * 8)
                record.append(cell(reason, style))
                continue

            lead_text = offer.get('leadTimeText')
            if offer.get('stockSufficient') is True:
                lead_text = 'In stock'
            record.extend([
                cell(offer.get('supplierPartNumber')),
                cell(offer.get('stock'), STYLE_INT),
                cell(lead_text),
                cell(offer.get('leadTimeDays'), STYLE_INT),
                cell(offer.get('unitPrice'), STYLE_MONEY_FINE),
                cell(offer.get('extendedPrice'), STYLE_MONEY),
                cell(offer.get('orderQuantity'), STYLE_INT),
                cell(offer.get('lifecycle'), SEVERITY_STYLE.get(offer.get('lifecycleSeverity'), STYLE_DEFAULT)),
                cell('Found'),
            ])

        record.extend([
            cell(comparison.get('bestPriceSupplier')),
            # Several suppliers can be equally fast, so all of them are listed.
            cell(' / '.join(comparison.get('bestLeadTimeSuppliers') or [])),
            cell(comparison.get('bestLeadTimeDays'), STYLE_INT),
            cell(comparison.get('recommendedSupplier')),
            cell(comparison.get('lifecycle'),
                 SEVERITY_STYLE.get(comparison.get('lifecycleSeverity'), STYLE_DEFAULT)),
            cell(leadtime_module.BAND_LABEL[bands[index]]),
            cell('; '.join(f['text'] for f in comparison.get('flags') or [])),
        ])
        rows.append(record)

    return rows


DISTRIBUTOR_COLUMNS = [
    'Row', 'Part Number', 'Quantity', 'Via', 'Distributor', 'Distributor P/N',
    'Packaging', 'Stock', 'Availability', 'Min Order', 'Order Multiple',
    'Order Qty', 'Unit Price', 'Extended Price', 'Currency', 'Covers Quantity',
]
DISTRIBUTOR_WIDTHS = [6, 26, 9, 14, 22, 22, 16, 11, 18, 10, 14, 10, 12, 14, 9, 15]


def has_distributor_detail(result):
    """True when any supplier returned a per-distributor breakdown."""
    for row in result['rows']:
        for offer in row['offers'].values():
            if offer and offer.get('distributorOffers'):
                return True
    return False


def attribution_notes(result):
    """Attribution lines required by any supplier that asks for them."""
    notes = []
    seen = set()
    for row in result['rows']:
        for offer in row['offers'].values():
            attribution = offer and offer.get('attribution')
            if not attribution:
                continue
            key = attribution.get('name')
            if key in seen:
                continue
            seen.add(key)
            notes.append('%s %s — %s' % (
                attribution.get('text') or 'Powered by',
                attribution.get('name') or '',
                attribution.get('home') or attribution.get('url') or '',
            ))
    return notes


def build_distributor_rows(result, summary, styled=False):
    """One row per distributor offer, for aggregators like TrustedParts.

    A single row of the main sheet can hide a dozen distributors; this is where
    they all become visible, filterable and sortable.
    """
    def cell(value, style=STYLE_DEFAULT):
        return Cell(value, style) if styled else value

    rows = [[cell(name, STYLE_HEADER) for name in DISTRIBUTOR_COLUMNS]]

    for row in result['rows']:
        for supplier_id, offer in row['offers'].items():
            if not offer or not offer.get('distributorOffers'):
                continue
            for entry in offer['distributorOffers']:
                covers = entry.get('stockSufficient')
                rows.append([
                    cell(row.get('row'), STYLE_INT),
                    cell(row.get('mpn')),
                    cell(row.get('quantity'), STYLE_INT),
                    cell(offer.get('supplier')),
                    cell(entry.get('distributor')),
                    cell(entry.get('supplierPartNumber')),
                    cell(entry.get('packaging')),
                    cell(entry.get('stock'), STYLE_INT),
                    cell(entry.get('availabilityText')),
                    cell(entry.get('minimumOrderQuantity'), STYLE_INT),
                    cell(entry.get('orderMultiple'), STYLE_INT),
                    cell(entry.get('orderQuantity'), STYLE_INT),
                    cell(entry.get('unitPrice'), STYLE_MONEY_FINE),
                    cell(entry.get('extendedPrice'), STYLE_MONEY),
                    cell(entry.get('currency')),
                    cell('yes' if covers else ('no' if covers is False else ''),
                         STYLE_DEFAULT if covers else STYLE_WARN),
                ])

    # Required attribution, in the first column of a trailing row.
    for note in attribution_notes(result):
        rows.append([cell(note, STYLE_MUTED)] + [cell('')] * (len(DISTRIBUTOR_COLUMNS) - 1))
    return rows


def write_csv(path, result, summary):
    rows = build_rows(result, summary, styled=False)
    with open(path, 'w', encoding='utf-8-sig', newline='') as handle:
        writer = csv.writer(handle)
        for row in rows:
            writer.writerow([_csv_cell(value) for value in row])

    base, extension = os.path.splitext(path)

    def companion(suffix, companion_rows):
        target = base + suffix + (extension or '.csv')
        with open(target, 'w', encoding='utf-8-sig', newline='') as handle:
            writer = csv.writer(handle)
            for row in companion_rows:
                writer.writerow([_csv_cell(value) for value in row])

    # CSV has no second sheet, so the extra tables go beside it as files.
    companion('-summary', build_parts_rows(result, summary))
    if has_distributor_detail(result):
        companion('-distributors', build_distributor_rows(result, summary))
    if result.get('excluded'):
        companion('-skipped', build_excluded_rows(result['excluded']))
    return path


def _csv_cell(value):
    if value is None:
        return ''
    text = str(value)
    # A leading =, +, - or @ makes a spreadsheet treat the cell as a formula.
    if text[:1] in ('=', '+', '-', '@') and not _is_number(text):
        return "'" + text
    return text


def _is_number(text):
    try:
        float(text)
        return True
    except ValueError:
        return False


def write_workbook(path, result, summary, meta=None):
    sheets = build_workbook_sheets(result, summary, meta)
    return write_xlsx(path, sheets, freeze_rows=1, autofilter=True)


def write_report_workbook(target, books):
    """One workbook covering one or more analyzed BOMs.

    `books` is a list of {'result', 'summary', 'meta', 'excluded'}. With a
    single BOM the sheets keep their plain names; with several, each set is
    prefixed with the BOM's name so the tabs stay readable.
    """
    books = list(books or [])
    sheets = []
    for index, book in enumerate(books):
        prefix = ''
        if len(books) > 1:
            meta = book.get('meta') or {}
            # Four sheet names share the prefix, and Excel truncates at 31
            # characters, so the BOM name gets the room that leaves.
            prefix = sheet_prefix(meta.get('name') or 'BOM %d' % (index + 1))
        sheets.extend(build_workbook_sheets(
            book['result'], book['summary'], book.get('meta'), book.get('excluded'), prefix,
        ))
    return write_xlsx(target, sheets, freeze_rows=1, autofilter=True)


def sheet_prefix(name):
    # "Full comparison" is the longest sheet name at 15 characters; 31 minus
    # that and a space leaves 15 for the BOM.
    cleaned = re.sub(r'[\[\]:*?/\\]', '-', str(name or '')).strip()
    return cleaned[:15].strip()


def write_json(path, result, summary):
    with open(path, 'w', encoding='utf-8') as handle:
        json.dump({
            'suppliers': result['suppliers'],
            'stats': result['stats'],
            'summary': summary,
            'rows': result['rows'],
            'excluded': result.get('excluded') or [],
        }, handle, indent=2)
    return path


WRITERS = {'csv': write_csv, 'xlsx': write_workbook, 'json': write_json}


# ── Terminal rendering ──────────────────────────────────────────────────────

ANSI = {
    'reset': '\033[0m', 'bold': '\033[1m', 'dim': '\033[2m',
    'red': '\033[31m', 'yellow': '\033[33m', 'green': '\033[32m',
    'cyan': '\033[36m', 'blue': '\033[34m',
}


class Palette:
    """ANSI colours, or no-ops when the output is piped or NO_COLOR is set."""

    def __init__(self, enabled=True):
        self.enabled = enabled

    def __call__(self, text, *styles):
        if not self.enabled or not styles:
            return str(text)
        prefix = ''.join(ANSI.get(s, '') for s in styles)
        return '%s%s%s' % (prefix, text, ANSI['reset'])


def _display_width(text):
    # Strip ANSI so padding is computed on what the terminal actually shows.
    return len(re.sub(r'\033\[[0-9;]*m', '', str(text)))


def render_table(headers, rows, aligns=None, palette=None):
    """A compact fixed-width table. Cells may already contain ANSI codes."""
    paint = palette or Palette(False)
    aligns = aligns or ['left'] * len(headers)
    widths = [_display_width(h) for h in headers]
    for row in rows:
        for i, value in enumerate(row):
            widths[i] = max(widths[i], _display_width(value))

    def pad(value, width, align):
        gap = width - _display_width(value)
        if gap <= 0:
            return str(value)
        return (' ' * gap + str(value)) if align == 'right' else (str(value) + ' ' * gap)

    lines = ['  '.join(paint(pad(h, widths[i], aligns[i]), 'bold') for i, h in enumerate(headers))]
    lines.append(paint('  '.join('─' * w for w in widths), 'dim'))
    for row in rows:
        lines.append('  '.join(pad(v, widths[i], aligns[i]) for i, v in enumerate(row)))
    return '\n'.join(lines)


def truncate(text, limit):
    text = '' if text is None else str(text)
    return text if len(text) <= limit else text[:limit - 1] + '…'


def money(value, currency='USD'):
    if value is None:
        return '—'
    symbol = {'USD': '$', 'EUR': '€', 'GBP': '£'}.get(currency, '')
    if symbol:
        return '%s%s' % (symbol, format(value, ',.2f') if value >= 1 else format(value, ',.4f'))
    return '%s %s' % (format(value, ',.2f'), currency)


def integer(value):
    return '—' if value is None else format(int(value), ',')


def lead_label(offer):
    if not offer or not offer.get('found'):
        return '—'
    if offer.get('stockSufficient') is True:
        return 'in stock'
    return offer.get('leadTimeText') or '—'


def format_days(days):
    return '—' if days is None else (format_lead_time(days) or str(days))


# ── Concise report ──────────────────────────────────────────────────────────
#
# The full comparison sheet is the audit trail: every supplier, every column.
# This is the sheet somebody actually reads — the headline numbers, who to buy
# from, and what needs a decision, in that order.

REPORT_WIDTH = 8

PARTS_COLUMNS = [
    'Row', 'Part Number', 'Qty', 'Manufacturer', 'Description',
    'Buy From', 'Unit Price', 'Extended', 'Lead Time', 'Lifecycle', 'Notes',
]
PARTS_WIDTHS = [6, 26, 8, 18, 36, 15, 12, 13, 15, 16, 46]

EXCLUDED_COLUMNS = ['Row', 'Part Number', 'Qty', 'Reference', 'Description', 'Why it was skipped']
EXCLUDED_WIDTHS = [6, 26, 8, 18, 36, 44]

ALTERNATE_COLUMNS = [
    'Row', 'Primary Part', 'Primary Status', 'Alternate', 'Qty',
    'Available', 'Stock', 'Lifecycle', 'Best Price', 'From', 'Notes',
]
ALTERNATE_WIDTHS = [6, 26, 20, 26, 8, 11, 13, 20, 13, 15, 40]

EXCLUSION_LABEL = {
    FLAGGED: 'Marked skip to production on the BOM',
    IGNORED: 'In-house part number',
    MERGED: 'Duplicate line, quantities added',
    DUPLICATE: 'Already covered by another BOM',
}


def _pad(row, width=REPORT_WIDTH):
    return row + [Cell('')] * max(0, width - len(row))


def _section(title, width=REPORT_WIDTH):
    return _pad([Cell(title, STYLE_SECTION)], width)


def risk_rows(result):
    """The lines a buyer has to make a decision about.

    A line earns its place by carrying a flag at bad or warn level — the same
    flags shown beside it — rather than by a rule restated here that could
    drift from them. An info flag (a split, a price spread) is worth reading
    but is not a decision.
    """
    risky = []
    for row in result['rows']:
        comparison = row['comparison']
        found = any(o and o.get('found') for o in row['offers'].values())
        if not found or any(f.get('level') in ('bad', 'warn')
                            for f in comparison.get('flags') or []):
            risky.append(row)
    return risky


def report_stats(result, summary, excluded=None):
    """The handful of numbers the top of the report leads with."""
    rows = result['rows']
    # Short means the suppliers' shelves together cannot cover it, not that
    # no single one can: a line coverable by splitting the order is not a risk.
    stock_risk = sum(1 for r in rows if not r['comparison'].get('stockCovers'))
    lifecycle_risk = sum(
        1 for r in rows if r['comparison'].get('lifecycleSeverity') in ('bad', 'warn')
    )
    return {
        'lines': summary.get('lines') or len(rows),
        'totalQuantity': summary.get('totalQuantity') or 0,
        'bestMixTotal': summary.get('bestMixTotal'),
        'stockRisk': stock_risk,
        'lifecycleRisk': lifecycle_risk,
        'notFound': summary.get('notFoundLines') or 0,
        'skipped': len(excluded or []),
    }


def build_report_rows(result, summary, meta=None, excluded=None):
    """The headline sheet: title block, KPI strip, supplier carts, decisions."""
    meta = meta or {}
    currency = summary.get('currency') or 'USD'
    suppliers = result['suppliers']
    stats = report_stats(result, summary, excluded)
    generated = meta.get('generated') or datetime.datetime.now().strftime('%Y-%m-%d %H:%M')

    rows = [
        _pad([Cell('BOM Supplier Report', STYLE_TITLE)]),
        _pad([Cell('%s · %s · prices in %s' % (
            meta.get('name') or 'Bill of materials', generated, currency), STYLE_SUBTITLE)]),
        _pad([]),
        _section('Overview'),
        _pad([Cell(label, STYLE_LABEL) for label in [
            'Lines', 'Units', 'Best-mix total', 'Stock risk',
            'Lifecycle risk', 'Not found', 'Skipped', 'Suppliers',
        ]]),
        _pad([
            Cell(stats['lines'], STYLE_INT_BOLD),
            Cell(stats['totalQuantity'], STYLE_INT_BOLD),
            Cell(stats['bestMixTotal'], STYLE_MONEY_BOLD),
            Cell(stats['stockRisk'], STYLE_BAD if stats['stockRisk'] else STYLE_GOOD),
            Cell(stats['lifecycleRisk'], STYLE_WARN if stats['lifecycleRisk'] else STYLE_GOOD),
            Cell(stats['notFound'], STYLE_BAD if stats['notFound'] else STYLE_GOOD),
            Cell(stats['skipped'], STYLE_MUTED),
            Cell(len(suppliers), STYLE_INT),
        ]),
        _pad([]),
        _section('What each supplier would cost'),
        _pad([Cell(label, STYLE_HEADER) for label in [
            'Supplier', 'Lines quoted', 'Not carried', 'Cannot cover alone', 'Cart total', '', '', '',
        ]]),
    ]

    totals = summary.get('supplierTotals') or {}
    cheapest = summary.get('cheapestSingleSource')
    for supplier in suppliers:
        entry = totals.get(supplier['id'])
        if not entry:
            continue
        name = supplier['name'] + (' ← cheapest single source' if supplier['id'] == cheapest else '')
        rows.append(_pad([
            Cell(name, STYLE_BOLD if supplier['id'] == cheapest else STYLE_DEFAULT),
            Cell(entry.get('linesPriced'), STYLE_INT),
            Cell(entry.get('linesMissing'), STYLE_INT),
            Cell(entry.get('linesShort'), STYLE_INT),
            Cell(entry.get('total'), STYLE_MONEY),
        ]))

    if len(suppliers) > 1:
        savings = summary.get('mixSavings')
        rows.append(_pad([
            Cell('Cheapest line by line', STYLE_BOLD),
            Cell(summary.get('bestMixLines'), STYLE_INT),
            Cell(''), Cell(''),
            Cell(summary.get('bestMixTotal'), STYLE_MONEY_BOLD),
            Cell('saves %s vs. single-sourcing' % _plain_money(savings, currency)
                 if isinstance(savings, (int, float)) and savings > 0 else '', STYLE_MUTED),
        ]))

    # Shaded by availability, the same four colours as everywhere else, so the
    # worst lines are findable by eye before anything is read.
    bands = leadtime_module.bands_by_row(result)
    band_by_row = {id(row): bands[index] for index, row in enumerate(result['rows'])}

    risky = risk_rows(result)
    rows.append(_pad([]))
    rows.append(_section('Needs a decision (%d)' % len(risky)))
    if risky:
        rows.append(_pad([Cell(label, STYLE_HEADER) for label in [
            'Row', 'Part Number', 'Qty', 'Availability', 'Lifecycle', 'Issue', '', '',
        ]]))
        for row in risky:
            comparison = row['comparison']
            cell = banded_cell(LEAD_FILL.get(band_by_row.get(id(row))))
            entry = leadtime_module.summarize_row(row, suppliers)
            rows.append(_pad([
                cell(row.get('row'), STYLE_INT),
                cell(row.get('mpn')),
                cell(row.get('quantity'), STYLE_INT),
                cell(entry['availability']),
                cell(comparison.get('lifecycle')),
                cell('; '.join(f['text'] for f in comparison.get('flags') or [])),
            ]))
        rows.append(_pad([]))
        rows.extend(availability_legend())
    else:
        rows.append(_pad([Cell('Every line is in stock, priced and in production.', STYLE_GOOD)]))

    counts = {}
    for entry in excluded or []:
        counts[entry.get('reason')] = counts.get(entry.get('reason'), 0) + 1
    if counts:
        rows.append(_pad([]))
        rows.append(_section('Lines not looked up'))
        for reason, label in EXCLUSION_LABEL.items():
            if counts.get(reason):
                rows.append(_pad([
                    Cell(label), Cell(counts[reason], STYLE_INT),
                    Cell('listed in full on the Skipped sheet', STYLE_MUTED),
                ]))

    notes = attribution_notes(result)
    if notes:
        rows.append(_pad([]))
        for note in notes:
            rows.append(_pad([Cell(note, STYLE_MUTED)]))

    return rows


def _plain_money(value, currency='USD'):
    return money(value, currency) if value is not None else '—'


def _also_in(row):
    """What other BOMs need of this part: [(name, quantity), ...].

    Only the browser knows which other BOMs are open, so it attaches this when
    it asks for a report. Absent, the column simply does not appear.
    """
    entries = row.get('alsoIn')
    if not isinstance(entries, list):
        return []
    usable = []
    for entry in entries:
        if isinstance(entry, dict) and entry.get('name'):
            usable.append((str(entry['name'])[:120], entry.get('quantity')))
    return usable


# Semantic style -> the kind of fill cell it becomes on a shaded row. A cell
# that was carrying meaning in its text colour (a lifecycle severity, a muted
# note) gives that up to the fill, which is the stronger signal of the two.
_FILL_KIND = {
    STYLE_INT: 'int',
    STYLE_INT_BOLD: 'int',
    STYLE_MONEY: 'money',
    STYLE_MONEY_BOLD: 'money',
    STYLE_MONEY_FINE: 'money_fine',
}


def banded_cell(colour):
    """A cell factory that shades everything it makes, or nothing if unshaded."""
    def cell(value, style=STYLE_DEFAULT):
        if colour is None:
            return Cell(value, style)
        return Cell(value, fill_style(colour, _FILL_KIND.get(style, 'text')))
    return cell


def availability_legend(width=REPORT_WIDTH):
    """The key to the colours, so a shaded sheet explains itself."""
    rows = [_pad([Cell('Colour key', STYLE_LABEL)], width)]
    for band in leadtime_module.BAND_ORDER:
        colour = LEAD_FILL.get(band)
        rows.append(_pad([
            Cell(leadtime_module.BAND_LABEL[band], fill_style(colour)),
            Cell(_BAND_EXPLANATION[band], fill_style(colour)),
        ], width))
    return rows


_BAND_EXPLANATION = {
    leadtime_module.NONE: 'No supplier searched can provide the part at this time',
    leadtime_module.LONG: 'More than %d weeks out' % (leadtime_module.LONG_DAYS // 7),
    leadtime_module.MEDIUM: '%d to %d weeks out' % (leadtime_module.MEDIUM_DAYS // 7,
                                                    leadtime_module.LONG_DAYS // 7),
    leadtime_module.UNKNOWN: 'Carried, but no supplier quoted a date',
    leadtime_module.QUICK: 'In stock, or quoted inside %d weeks' % (
        leadtime_module.MEDIUM_DAYS // 7),
}


def build_parts_rows(result, summary, styled=False):
    """One line per part, only the columns a buyer acts on.

    Shaded by how soon the part can arrive, the same four colours the lead-time
    report uses: a buyer opening the summary should not have to cross-reference
    another sheet to see which lines are the problem.
    """
    def cell(value, style=STYLE_DEFAULT):
        return Cell(value, style) if styled else value

    bands = leadtime_module.bands_by_row(result) if styled else []
    shared = any(_also_in(row) for row in result['rows'])
    alternates = has_alternates(result)
    columns = list(PARTS_COLUMNS)
    if shared:
        columns.insert(3, 'Also In')
    if alternates:
        columns.insert(-1, 'Approved Alternates')

    rows = [[cell(name, STYLE_HEADER) for name in columns]]

    for index, row in enumerate(result['rows']):
        comparison = row['comparison']
        if styled:
            cell = banded_cell(LEAD_FILL.get(bands[index]))
        # The recommended supplier is the one to price: it already balances
        # "soonest" against "cheapest among the soonest".
        chosen = None
        for offer in row['offers'].values():
            if offer and offer.get('found') and offer.get('supplier') == comparison.get('recommendedSupplier'):
                chosen = offer
                break

        lead = '—'
        if chosen:
            lead = 'In stock' if chosen.get('stockSufficient') is True else (chosen.get('leadTimeText') or '—')

        notes = [f['text'] for f in comparison.get('flags') or []]
        if chosen and chosen.get('aggregator') and chosen.get('distributor'):
            notes.insert(0, 'via %s' % chosen['distributor'])
        if row.get('mergedRows'):
            notes.append('includes rows %s' % ', '.join(str(r) for r in row['mergedRows']))

        record = [
            cell(row.get('row'), STYLE_INT),
            cell(row.get('mpn')),
            cell(row.get('quantity'), STYLE_INT),
        ]
        if shared:
            record.append(cell(', '.join(
                ('%s (%s)' % (name, qty)) if isinstance(qty, (int, float)) else name
                for name, qty in _also_in(row)
            ) or None, STYLE_MUTED))
        plan = comparison.get('allocation') or {}
        split = bool(plan.get('splitRequired')) and not plan.get('shortfall')
        record.extend([
            cell(row.get('manufacturer')),
            cell(row.get('description')),
            cell('%d suppliers, split' % plan['suppliers'] if split
                 else (comparison.get('recommendedSupplier') or '—')),
            cell(None if split else (chosen.get('unitPrice') if chosen else None), STYLE_MONEY_FINE),
            cell(plan.get('total') if split
                 else (chosen.get('extendedPrice') if chosen else None), STYLE_MONEY),
            cell('In stock, split' if split else lead),
            cell(comparison.get('lifecycle'),
                 SEVERITY_STYLE.get(comparison.get('lifecycleSeverity'), STYLE_DEFAULT)),
        ])
        if alternates:
            record.append(cell(alternate_summary(row), STYLE_MUTED))
        record.append(cell('; '.join(notes)))
        rows.append(record)

    if styled and rows[1:]:
        # The key goes under the table rather than over it, so the sheet still
        # opens on its first part and the header stays row 1. The caller keeps
        # it out of the filter range — see parts_filter_rows().
        width = len(rows[0])
        rows.append([Cell('')] * width)
        for entry in availability_legend(width):
            rows.append(list(entry))
    return rows


def parts_filter_rows(result):
    """How many rows of the Parts sheet are the table, header included."""
    return 1 + len(result['rows'])


def has_alternates(result):
    """True when any analyzed line named an approved alternate."""
    return any(row.get('alternates') for row in result['rows'])


def alternate_summary(row):
    """One cell's worth: what the BOM offers instead, and whether it is there."""
    entries = [a for a in (row.get('alternates') or []) if a.get('mpn')]
    if not entries:
        return None
    return '; '.join(
        '%s (%s)' % (entry['mpn'], 'available' if entry.get('usable') else 'not available')
        for entry in entries
    )


def build_alternate_rows(result, styled=False):
    """One row per approved alternate, with the primary it stands in for.

    Separate from the parts sheet because the question is different: not what
    to buy, but what could be bought instead, and how ready that answer is.
    """
    def cell(value, style=STYLE_DEFAULT):
        return Cell(value, style) if styled else value

    rows = [[cell(name, STYLE_HEADER) for name in ALTERNATE_COLUMNS]]

    for row in result['rows']:
        comparison = row.get('comparison') or {}
        for entry in row.get('alternates') or []:
            if not entry.get('mpn'):
                continue
            notes = []
            if not entry.get('found'):
                notes.append('No supplier carries it')
            elif not entry.get('coversQuantity'):
                notes.append('Nobody holds the full quantity')
            if entry.get('lifecycleSeverity') == 'bad':
                notes.append('Alternate is itself ending')

            rows.append([
                cell(row.get('row'), STYLE_INT),
                cell(row.get('mpn')),
                cell(comparison.get('lifecycle'),
                     SEVERITY_STYLE.get(comparison.get('lifecycleSeverity'), STYLE_DEFAULT)),
                cell(entry.get('mpn')),
                cell(entry.get('quantity'), STYLE_INT),
                cell('yes' if entry.get('usable') else 'no',
                     STYLE_GOOD if entry.get('usable') else STYLE_WARN),
                cell(entry.get('stock'), STYLE_INT),
                cell(entry.get('lifecycle'),
                     SEVERITY_STYLE.get(entry.get('lifecycleSeverity'), STYLE_DEFAULT)),
                cell(entry.get('bestPrice'), STYLE_MONEY),
                cell(entry.get('bestPriceSupplier')),
                cell('; '.join(notes)),
            ])
    return rows


def build_excluded_rows(excluded, styled=False):
    def cell(value, style=STYLE_DEFAULT):
        return Cell(value, style) if styled else value

    rows = [[cell(name, STYLE_HEADER) for name in EXCLUDED_COLUMNS]]
    for entry in excluded or []:
        rows.append([
            cell(entry.get('row'), STYLE_INT),
            cell(entry.get('mpn')),
            cell(entry.get('quantity'), STYLE_INT),
            cell(entry.get('reference')),
            cell(entry.get('description')),
            cell(entry.get('detail') or EXCLUSION_LABEL.get(entry.get('reason'), ''), STYLE_MUTED),
        ])
    return rows


def _parts_widths(header):
    """Widths for the parts sheet, which grows a column or two conditionally."""
    extra = {'Also In': 24, 'Approved Alternates': 34}
    widths = []
    base = dict(zip(PARTS_COLUMNS, PARTS_WIDTHS))
    for name in header:
        label = getattr(name, 'value', name)
        widths.append(base.get(label, extra.get(label, 18)))
    return widths


# ── Split orders ────────────────────────────────────────────────────────────

SPLIT_COLUMNS = [
    'Row', 'Part Number', 'Needed', 'Supplier', 'Supplier P/N',
    'Take', 'Order Qty', 'They Hold', 'Unit Price', 'Extended', 'Notes',
]
SPLIT_WIDTHS = [6, 26, 10, 20, 24, 10, 11, 12, 12, 13, 44]


def split_lines(result):
    """Every line whose quantity has to come from more than one supplier."""
    out = []
    for row in result['rows']:
        plan = (row.get('comparison') or {}).get('allocation') or {}
        if plan.get('splitRequired'):
            out.append((row, plan))
    return out


def has_split_orders(result):
    return bool(split_lines(result))


def build_split_rows(result, styled=False):
    """One row per purchase order, grouped under the BOM line that needs them."""
    def cell(value, style=STYLE_DEFAULT):
        return Cell(value, style) if styled else value

    rows = [[cell(name, STYLE_HEADER) for name in SPLIT_COLUMNS]]
    for row, plan in split_lines(result):
        for index, line in enumerate(plan['lines']):
            note = ''
            if index == 0:
                note = ('%d of %d covered — %d still to find'
                        % (plan['covered'], plan['needed'], plan['shortfall'])
                        if plan['shortfall'] else
                        'Covered in full across %d purchase orders' % plan['suppliers'])
                if plan.get('backorder'):
                    note += '; soonest for the rest is %s in %s' % (
                        plan['backorder']['supplier'],
                        format_lead_time(plan['backorder']['leadTimeDays']))
            rows.append([
                cell(row.get('row') if index == 0 else None, STYLE_INT),
                cell(row.get('mpn') if index == 0 else None),
                cell(plan['needed'] if index == 0 else None, STYLE_INT),
                cell(line['supplier']),
                cell(line.get('supplierPartNumber')),
                cell(line['take'], STYLE_INT),
                cell(line['orderQuantity'], STYLE_INT),
                cell(line['stock'], STYLE_INT),
                cell(line.get('unitPrice'), STYLE_MONEY_FINE),
                cell(line.get('extendedPrice'), STYLE_MONEY),
                cell(note, STYLE_MUTED),
            ])
    return rows


# ── Long lead times ─────────────────────────────────────────────────────────

LEAD_COLUMNS = [
    'Row', 'Part Number', 'Qty', 'Manufacturer', 'Description',
    'Availability', 'Lead (days)', 'Buy From', 'Supplier P/N',
    'Unit Price', 'Extended', 'Stock', 'Lifecycle',
    'Suppliers Carrying', 'Also Available From', 'Notes',
]
LEAD_WIDTHS = [6, 26, 8, 18, 34, 20, 11, 16, 22, 12, 13, 12, 15, 10, 40, 46]

# The colour each band is shaded, which is also what the legend explains.
# Defined once and used by every sheet that shades a part, so the summary
# report and the lead-time report can never drift on to different palettes.
LEAD_FILL = {
    leadtime_module.QUICK: 'green',
    leadtime_module.MEDIUM: 'yellow',
    leadtime_module.LONG: 'orange',
    leadtime_module.NONE: 'red',
    # Deliberately unshaded: "nobody would say when" is not a duration, and
    # colouring it as one would assert something no supplier did.
    leadtime_module.UNKNOWN: None,
}


def _others_text(entry):
    return '; '.join(
        '%s (%s)' % (other['supplier'], other['availability'])
        for other in entry['others'][:4]
    )


def build_lead_rows(report, styled=False):
    """The lead-time table, one row per part, shaded by band."""
    header = list(LEAD_COLUMNS)
    rows = [[Cell(h, STYLE_HEADER) for h in header] if styled else header]

    for entry in report['rows']:
        colour = LEAD_FILL.get(entry['band'])
        note = entry['note']

        values = [
            entry['row'], entry['mpn'], entry['quantity'], entry['manufacturer'],
            entry['description'], entry['availability'], entry['days'],
            entry['supplier'], entry['supplierPartNumber'],
            entry['unitPrice'], entry['extendedPrice'], entry['stock'],
            entry['lifecycle'], entry['suppliersCarrying'],
            _others_text(entry), note,
        ]
        if not styled:
            rows.append(values)
            continue

        # Which cells are numbers decides the format; the band decides the fill.
        kinds = ['int', 'text', 'int', 'text', 'text', 'text', 'int', 'text',
                 'text', 'money', 'money', 'int', 'text', 'int', 'text', 'text']
        rows.append([
            Cell(value, fill_style(colour, kind))
            for value, kind in zip(values, kinds)
        ])
    return rows


def build_lead_summary_rows(report, meta=None):
    """The title block above the table: what was searched, and the legend."""
    meta = meta or {}
    counts = report['counts']
    thresholds = report['thresholds']

    rows = [
        _pad([Cell('Long Lead Times', STYLE_TITLE)]),
        _pad([Cell('%s · %s · %d part%s across %d supplier%s' % (
            meta.get('name') or 'Parts looked up',
            meta.get('generated') or '',
            len(report['rows']), '' if len(report['rows']) == 1 else 's',
            len(report['suppliers']), '' if len(report['suppliers']) == 1 else 's',
        ), STYLE_SUBTITLE)]),
        _pad([]),
        _section('How soon each part can arrive'),
    ]

    legend = [
        (leadtime_module.NONE, 'red',
         'No supplier searched can provide the part at this time'),
        (leadtime_module.LONG, 'orange',
         'More than %d weeks out' % (thresholds['longDays'] // 7)),
        (leadtime_module.MEDIUM, 'yellow',
         '%d to %d weeks out' % (thresholds['mediumDays'] // 7,
                                 thresholds['longDays'] // 7)),
        (leadtime_module.QUICK, 'green',
         'In stock, or quoted inside %d weeks' % (thresholds['mediumDays'] // 7)),
        (leadtime_module.UNKNOWN, None,
         'Carried, but no supplier quoted a date'),
    ]
    for band, colour, explanation in legend:
        rows.append(_pad([
            Cell(leadtime_module.BAND_LABEL[band], fill_style(colour)),
            Cell(counts.get(band, 0), fill_style(colour, 'int')),
            Cell(explanation, fill_style(colour)),
        ]))

    rows.append(_pad([]))
    rows.append(_pad([Cell(
        'Stock on hand counts as zero days, because it ships today whatever the '
        'factory quotes behind it. Where several suppliers can deliver in the same '
        'time, the cheapest of them is the one named.', STYLE_MUTED)]))
    return rows


def build_lead_workbook_sheets(report, meta=None, prefix=''):
    """A lead-time workbook: the legend, then the table."""
    def name(base):
        return ('%s %s' % (prefix, base)).strip() if prefix else base

    summary_rows = build_lead_summary_rows(report, meta)
    return [{
        'name': name('Lead times'),
        'rows': summary_rows,
        'widths': [30, 10, 60, 15, 15, 15, 15, 15],
        'freeze': 0,
        'autofilter': False,
        'merges': ['A1:%s1' % column_letter(REPORT_WIDTH - 1),
                   'A2:%s2' % column_letter(REPORT_WIDTH - 1)],
        'heights': {0: 26, 1: 18},
    }, {
        'name': name('By part'),
        'rows': build_lead_rows(report, styled=True),
        'widths': LEAD_WIDTHS,
    }]


def write_lead_workbook(target, report, meta=None):
    write_xlsx(target, build_lead_workbook_sheets(report, meta))


def build_workbook_sheets(result, summary, meta=None, excluded=None, prefix=''):
    """Every sheet of the deliverable, headline first.

    `prefix` names the sheets when several BOMs share one workbook. Excel caps
    a sheet name at 31 characters, which the writer enforces.
    """
    excluded = excluded if excluded is not None else (result.get('excluded') or [])

    def name(base):
        return ('%s %s' % (prefix, base)).strip() if prefix else base

    report_rows = build_report_rows(result, summary, meta, excluded)
    parts_rows = build_parts_rows(result, summary, styled=True)
    sheets = [{
        'name': name('Report'),
        'rows': report_rows,
        'widths': [22, 15, 15, 15, 15, 15, 15, 15],
        # A title block is not a header row: freezing or filtering it would
        # only get in the way.
        'freeze': 0,
        'autofilter': False,
        'merges': ['A1:%s1' % column_letter(REPORT_WIDTH - 1),
                   'A2:%s2' % column_letter(REPORT_WIDTH - 1)],
        'heights': {0: 26, 1: 18},
    }, {
        'name': name('Parts'),
        'rows': parts_rows,
        'widths': PARTS_WIDTHS if len(parts_rows[0]) == len(PARTS_COLUMNS)
                  else _parts_widths(parts_rows[0]),
        'filterRows': parts_filter_rows(result),
    }, {
        'name': name('Full comparison'),
        'rows': build_rows(result, summary, styled=True),
        'widths': column_widths(result['suppliers']),
    }]

    if has_split_orders(result):
        sheets.append({
            'name': name('Split orders'),
            'rows': build_split_rows(result, styled=True),
            'widths': SPLIT_WIDTHS,
        })

    sheets.append({
        'name': name('Lead times'),
        'rows': build_lead_rows(leadtime_module.build_report(result), styled=True),
        'widths': LEAD_WIDTHS,
    })

    if has_distributor_detail(result):
        sheets.append({
            'name': name('Distributors'),
            'rows': build_distributor_rows(result, summary, styled=True),
            'widths': DISTRIBUTOR_WIDTHS,
        })
    if has_alternates(result):
        sheets.append({
            'name': name('Alternates'),
            'rows': build_alternate_rows(result, styled=True),
            'widths': ALTERNATE_WIDTHS,
        })
    if excluded:
        sheets.append({
            'name': name('Skipped'),
            'rows': build_excluded_rows(excluded, styled=True),
            'widths': EXCLUDED_WIDTHS,
        })
    return sheets
