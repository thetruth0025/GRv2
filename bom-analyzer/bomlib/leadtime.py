"""Who can supply each part soonest, and how long the slowest ones will take.

A purchasing question rather than a pricing one: not "what does this cost" but
"when can I have it, and from whom". Stock on hand is the fastest answer there
is, so a supplier holding enough ships at zero days regardless of what the
factory quotes behind it. When several suppliers can deliver in the same time,
the cheapest of them wins — being equally fast, price is the only thing left to
choose on.

Every part that was looked up appears, banded by how long it takes, so the
report reads as a worklist: the parts nobody carries first, then the long ones,
then everything that is simply fine.
"""

from .normalize import format_lead_time

# Three weeks and eight weeks, in days. A part quoted inside three weeks is
# treated as quick alongside stock on hand: the band exists to separate "order
# it" from "plan around it", and a fortnight is not something to plan around.
MEDIUM_DAYS = 21
LONG_DAYS = 56

# Bands, worst first — which is also the order the report sorts in.
NONE = 'none'          # nobody searched carries it
LONG = 'long'          # more than eight weeks
MEDIUM = 'medium'      # three to eight weeks
UNKNOWN = 'unknown'    # carried, but no supplier would say when
QUICK = 'quick'        # in stock, or inside three weeks

BAND_ORDER = (NONE, LONG, MEDIUM, UNKNOWN, QUICK)

BAND_LABEL = {
    NONE: 'Not available',
    LONG: 'Over 8 weeks',
    MEDIUM: '3–8 weeks',
    UNKNOWN: 'Unknown',
    QUICK: 'In stock or under 3 weeks',
}

# The short form goes in the availability column, which has to stay one line
# beside nine others; the full sentence goes in the note, where there is room
# to say it properly.
NO_SUPPLIER_SHORT = 'Not available'
NO_SUPPLIER_TEXT = 'No supplier searched can provide this part at this time'
UNKNOWN_SHORT = 'No date quoted'
UNKNOWN_TEXT = 'Carried, but no supplier quoted a date'


def effective_days(offer):
    """Days until it can ship, or None when nobody would say.

    Enough stock on hand is zero days: it goes out today, whatever the factory
    quotes for the next batch.
    """
    if not offer or not offer.get('found'):
        return None
    if offer.get('stockSufficient') is True:
        return 0
    days = offer.get('leadTimeDays')
    return days if isinstance(days, (int, float)) else None


def _price_key(offer):
    price = offer.get('extendedPrice')
    return price if isinstance(price, (int, float)) else float('inf')


def rank_offers(row, suppliers):
    """Every supplier that carries the part, soonest first, cheapest to break ties.

    Offers that carry the part but name no date sort last: they are real supply
    and belong in the report, but they cannot win a race they never entered.
    """
    entries = []
    for supplier in suppliers:
        offer = (row.get('offers') or {}).get(supplier['id'])
        if not offer or not offer.get('found'):
            continue
        days = effective_days(offer)
        entries.append({
            'supplier': offer.get('supplier') or supplier['name'],
            'supplierId': supplier['id'],
            'offer': offer,
            'days': days,
        })

    entries.sort(key=lambda e: (
        0 if e['days'] is not None else 1,
        e['days'] if e['days'] is not None else 0,
        _price_key(e['offer']),
        e['supplier'],
    ))
    return entries


def band_for(days, carried):
    if not carried:
        return NONE
    if days is None:
        return UNKNOWN
    if days <= 0:
        return QUICK
    if days < MEDIUM_DAYS:
        return QUICK
    if days <= LONG_DAYS:
        return MEDIUM
    return LONG


def availability_text(entry):
    """What a buyer would say out loud about this line."""
    if entry is None:
        return NO_SUPPLIER_SHORT
    days = entry['days']
    if days is None:
        return UNKNOWN_SHORT
    if days <= 0:
        return 'In stock'
    return format_lead_time(days)


def summarize_row(row, suppliers):
    """One line of the lead-time report."""
    ranked = rank_offers(row, suppliers)
    timed = [e for e in ranked if e['days'] is not None]
    winner = timed[0] if timed else (ranked[0] if ranked else None)
    days = winner['days'] if winner else None

    # A line nobody can cover alone but everybody can cover between them ships
    # today, from several purchase orders. Reporting it at the factory lead
    # time of whichever supplier happens to be quickest would be wrong twice
    # over: too slow, and from the wrong supplier.
    plan = (row.get('comparison') or {}).get('allocation') or {}
    split = bool(plan.get('splitRequired')) and not plan.get('shortfall')
    if split:
        days = plan.get('leadTimeDays', 0)

    band = band_for(days, bool(ranked))

    # Everyone else who could also supply it, so a second source is visible
    # without opening another report.
    others = []
    for entry in ranked:
        if winner is not None and entry is winner:
            continue
        others.append({
            'supplier': entry['supplier'],
            'days': entry['days'],
            'availability': availability_text(entry),
            'extendedPrice': entry['offer'].get('extendedPrice'),
            'unitPrice': entry['offer'].get('unitPrice'),
        })

    offer = winner['offer'] if winner else None
    tied = [] if split else ([e['supplier'] for e in timed if e['days'] == days] if timed else [])
    alternates = [
        {'mpn': a.get('mpn'), 'usable': a.get('usable')}
        for a in (row.get('alternates') or [])
        if isinstance(a, dict) and a.get('mpn')
    ]

    return {
        'index': row.get('index'),
        'row': row.get('row'),
        'mpn': row.get('mpn'),
        'quantity': row.get('quantity'),
        'manufacturer': row.get('manufacturer'),
        'description': row.get('description'),
        'reference': row.get('reference'),
        'band': band,
        'bandLabel': BAND_LABEL[band],
        'days': days,
        'availability': 'In stock, split' if split else availability_text(winner),
        # The purchase orders this line actually needs, so the report answers
        # "how do I get 200" and not only "who is quickest".
        'split': split,
        'allocation': plan.get('lines') or [],
        'supplier': ('%d suppliers, split' % plan['suppliers']) if split
                    else (winner['supplier'] if winner else None),
        'supplierPartNumber': None if split else (offer or {}).get('supplierPartNumber'),
        'unitPrice': None if split else (offer or {}).get('unitPrice'),
        'extendedPrice': plan.get('total') if split else (offer or {}).get('extendedPrice'),
        'orderQuantity': (offer or {}).get('orderQuantity'),
        'stock': plan.get('combinedStock') if split else (offer or {}).get('stock'),
        'lifecycle': (row.get('comparison') or {}).get('lifecycle'),
        'currency': (offer or {}).get('currency'),
        'suppliersCarrying': len(ranked),
        # Named only when the choice was actually made on price, so the report
        # can say why this supplier and not the other one.
        'tiedOn': tied if len(tied) > 1 else [],
        'others': others,
        'note': _note(band, winner, tied, ranked, alternates, split, plan),
        # A BOM alternate is the one thing that can rescue a line nobody
        # carries, so it rides along rather than living in another report.
        'alternates': alternates,
    }


def _note(band, winner, tied, ranked, alternates, split=False, plan=None):
    """The one sentence explaining this line, or nothing when it explains itself."""
    pieces = []
    plan = plan or {}
    if band == NONE:
        pieces.append(NO_SUPPLIER_TEXT)
    elif band == UNKNOWN:
        pieces.append('%d carr%s it, none quoting a date' % (
            len(ranked), 'ies' if len(ranked) == 1 else 'y'))
    elif len(tied) > 1:
        pieces.append('Cheapest of %d suppliers equally fast (%s)'
                      % (len(tied), ', '.join(tied)))
    elif split:
        pieces.append('Split %d across %d suppliers: %s' % (
            plan['needed'], plan['suppliers'],
            ', '.join('%s %d' % (line['supplier'], line['take']) for line in plan['lines'])))
    elif len(ranked) == 1 and band in (LONG, MEDIUM):
        pieces.append('Single source — only %s carries it' % winner['supplier'])

    # The one thing that can rescue a line that is late or gone: a substitute
    # somebody with the schematic already approved.
    if band in (NONE, LONG):
        usable = [a['mpn'] for a in alternates if a.get('usable')]
        if usable:
            pieces.append('BOM alternate available: %s' % ', '.join(usable[:2]))
    return '; '.join(pieces)


def bands_by_row(result):
    """The availability band of every line, keyed by its position in the result.

    One pass, so the summary report can shade its tables the same colours the
    lead-time report uses without recomputing the banding per sheet — and
    without any chance of the two disagreeing.
    """
    suppliers = result.get('suppliers') or []
    return [summarize_row(row, suppliers)['band'] for row in result.get('rows') or []]


def build_report(result, scope=None):
    """Every looked-up part, banded and sorted worst first."""
    suppliers = result.get('suppliers') or []
    rows = [summarize_row(row, suppliers) for row in result.get('rows') or []]
    order = {band: i for i, band in enumerate(BAND_ORDER)}
    rows.sort(key=lambda r: (
        order[r['band']],
        # Longest first inside a band: the report is read from the top.
        -(r['days'] if isinstance(r['days'], (int, float)) else 0),
        r['mpn'] or '',
    ))

    counts = {band: 0 for band in BAND_ORDER}
    for entry in rows:
        counts[entry['band']] += 1

    return {
        'scope': scope,
        'suppliers': suppliers,
        'rows': rows,
        'counts': counts,
        'thresholds': {'mediumDays': MEDIUM_DAYS, 'longDays': LONG_DAYS},
    }
