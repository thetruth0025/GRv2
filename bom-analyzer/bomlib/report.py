"""Turn analysis results into the tables the CLI writes and prints.

One column layout feeds the CSV writer, the .xlsx writer and the terminal
table, so the three can never drift apart.
"""

import csv
import json
import os
import re

from .normalize import format_lead_time
from .xlsx_writer import (
    Cell,
    STYLE_BAD,
    STYLE_DEFAULT,
    STYLE_HEADER,
    STYLE_INT,
    STYLE_MONEY,
    STYLE_MONEY_FINE,
    STYLE_MUTED,
    STYLE_WARN,
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
    'Recommended', 'Worst Lifecycle', 'Notes',
]
TAIL_WIDTHS = [17, 20, 14, 14, 18, 52]

SEVERITY_STYLE = {'bad': STYLE_BAD, 'warn': STYLE_WARN, 'ok': STYLE_DEFAULT, 'unknown': STYLE_MUTED}


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

    def cell(value, style=STYLE_DEFAULT):
        return Cell(value, style) if styled else value

    rows = [[cell(name, STYLE_HEADER) for name in build_header(suppliers, currency)]]

    for row in result['rows']:
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

    # CSV has no second sheet, so the distributor breakdown goes beside it.
    if has_distributor_detail(result):
        base, extension = os.path.splitext(path)
        companion = base + '-distributors' + (extension or '.csv')
        with open(companion, 'w', encoding='utf-8-sig', newline='') as handle:
            writer = csv.writer(handle)
            for row in build_distributor_rows(result, summary):
                writer.writerow([_csv_cell(value) for value in row])
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


def write_workbook(path, result, summary):
    sheets = [{
        'name': 'BOM Comparison',
        'rows': build_rows(result, summary, styled=True),
        'widths': column_widths(result['suppliers']),
    }]
    if has_distributor_detail(result):
        sheets.append({
            'name': 'Distributors',
            'rows': build_distributor_rows(result, summary, styled=True),
            'widths': DISTRIBUTOR_WIDTHS,
        })
    return write_xlsx(path, sheets, freeze_rows=1, autofilter=True)


def write_json(path, result, summary):
    with open(path, 'w', encoding='utf-8') as handle:
        json.dump({
            'suppliers': result['suppliers'],
            'stats': result['stats'],
            'summary': summary,
            'rows': result['rows'],
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
