"""Flatten each supplier's vocabulary into one comparable shape.

Every supplier reports lead time, lifecycle and pricing differently. Everything
here exists so the UI can put DigiKey and Mouser side by side.
"""

import math
import re


class Lifecycle:
    ACTIVE = 'Active'
    NRND = 'Not Recommended for New Designs'
    LAST_TIME_BUY = 'Last Time Buy'
    END_OF_LIFE = 'End of Life'
    OBSOLETE = 'Obsolete'
    DISCONTINUED = 'Discontinued'
    PREVIEW = 'Preview'
    NEW = 'New Product'
    UNKNOWN = 'Unknown'


# Lower rank is healthier. Used to pick the worst status across suppliers so a
# part flagged obsolete by either one is never shown as simply "Active".
LIFECYCLE_RANK = {
    Lifecycle.ACTIVE: 0,
    Lifecycle.NEW: 0,
    Lifecycle.PREVIEW: 1,
    Lifecycle.UNKNOWN: 2,
    Lifecycle.NRND: 3,
    Lifecycle.LAST_TIME_BUY: 4,
    Lifecycle.END_OF_LIFE: 5,
    Lifecycle.DISCONTINUED: 6,
    Lifecycle.OBSOLETE: 7,
}

LIFECYCLE_SEVERITY = {
    Lifecycle.ACTIVE: 'ok',
    Lifecycle.NEW: 'ok',
    Lifecycle.PREVIEW: 'info',
    Lifecycle.UNKNOWN: 'unknown',
    Lifecycle.NRND: 'warn',
    Lifecycle.LAST_TIME_BUY: 'warn',
    Lifecycle.END_OF_LIFE: 'bad',
    Lifecycle.DISCONTINUED: 'bad',
    Lifecycle.OBSOLETE: 'bad',
}


def round_to(value, decimals):
    """Round half-up, matching the arithmetic the frontend does on the same numbers."""
    factor = 10 ** decimals
    return math.floor(value * factor + 0.5) / factor


def normalize_lifecycle(raw):
    if raw is None:
        return Lifecycle.UNKNOWN
    text = str(raw).strip().lower()
    if not text:
        return Lifecycle.UNKNOWN

    if re.search(r'obsolete', text):
        return Lifecycle.OBSOLETE
    if re.search(r'last\s*time\s*buy|\bltb\b', text):
        return Lifecycle.LAST_TIME_BUY
    if re.search(r'end\s*of\s*life|\beol\b', text):
        return Lifecycle.END_OF_LIFE
    if re.search(r'discontinu', text):
        return Lifecycle.DISCONTINUED
    # DigiKey says "Not For New Designs", Mouser "Not Recommended for New Designs".
    if re.search(r'not\s*(recommended|for\s*new)|\bnrnd\b|no\s*longer\s*manufactured', text):
        return Lifecycle.NRND
    if re.search(r'preliminary|preview|pre-?release', text):
        return Lifecycle.PREVIEW
    if re.search(r'new\s*(product|at)', text):
        return Lifecycle.NEW
    if re.search(r'active|production|normally\s*stocking', text):
        return Lifecycle.ACTIVE
    return Lifecycle.UNKNOWN


def worst_lifecycle(statuses):
    worst = None
    for status in statuses:
        if not status or status not in LIFECYCLE_RANK:
            continue
        if worst is None or LIFECYCLE_RANK[status] > LIFECYCLE_RANK[worst]:
            worst = status
    return worst or Lifecycle.UNKNOWN


def lifecycle_severity(status):
    return LIFECYCLE_SEVERITY.get(status, 'unknown')


def parse_lead_time_days(raw):
    """Suppliers write lead time as prose: "12 Weeks", "45 Days", "In Stock"."""
    if raw is None:
        return None
    if isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)):
        return int(round_to(raw, 0)) if raw > 0 else None

    text = str(raw).strip().lower()
    if not text:
        return None
    if re.fullmatch(r'(in\s*stock|stock|immediate|available)', text):
        return 0

    match = re.search(r'(\d+(?:\.\d+)?)\s*([a-z]*)', text)
    if not match:
        return None
    try:
        value = float(match.group(1))
    except ValueError:
        return None
    if value <= 0:
        return None
    unit = match.group(2)

    if unit.startswith('w'):
        return int(round_to(value * 7, 0))
    if unit.startswith('m'):
        return int(round_to(value * 30, 0))
    if unit.startswith('y'):
        return int(round_to(value * 365, 0))
    if unit.startswith('d'):
        return int(round_to(value, 0))
    # A bare number in a lead-time field is conventionally weeks at both
    # suppliers, which is also the safer assumption to surface to a buyer.
    return int(round_to(value * 7, 0))


def format_lead_time(days):
    if days is None:
        return None
    if days == 0:
        return 'In stock'
    if days >= 7 and days % 7 == 0:
        weeks = days // 7
        return '1 week' if weeks == 1 else '%d weeks' % weeks
    return '1 day' if days == 1 else '%d days' % days


def parse_money(raw):
    """Mouser returns strings like "$1.23" or "1.234,56 €" depending on locale."""
    if raw is None or isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)):
        return float(raw)

    text = re.sub(r'[^0-9.,-]', '', str(raw).strip())
    if not text:
        return None
    last_comma = text.rfind(',')
    last_dot = text.rfind('.')
    if last_comma > last_dot:
        text = text.replace('.', '').replace(',', '.')
    else:
        text = text.replace(',', '')
    try:
        return float(text)
    except ValueError:
        return None


def parse_quantity(raw):
    if raw is None or isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)):
        return int(raw)
    text = re.sub(r'[^0-9-]', '', str(raw))
    if not text:
        return None
    try:
        return int(text)
    except ValueError:
        return None


def price_at_quantity(price_breaks, quantity):
    """The applicable break is the highest one at or below the quantity."""
    if not price_breaks:
        return None
    qty = quantity if isinstance(quantity, (int, float)) and quantity > 0 else 1
    usable = [
        b for b in price_breaks
        if b.get('quantity') is not None and b.get('unitPrice') is not None
    ]
    if not usable:
        return None
    usable = sorted(usable, key=lambda b: b['quantity'])

    chosen = None
    for brk in usable:
        if brk['quantity'] <= qty:
            chosen = brk
    # Below the smallest break the buyer still pays that break's unit price.
    return chosen if chosen else usable[0]


def order_quantity(needed, moq, multiple):
    """Suppliers sell in packaging multiples, so the purchased quantity can
    exceed what the BOM calls for."""
    qty = needed if isinstance(needed, (int, float)) and needed > 0 else 1
    minimum = moq if isinstance(moq, (int, float)) and moq > 0 else 1
    mult = multiple if isinstance(multiple, (int, float)) and multiple > 0 else 1
    order = max(qty, minimum)
    if mult > 1:
        order = int(math.ceil(order / mult) * mult)
    return int(order)


def pick_variation(variations, needed):
    """Choose the packaging option to actually buy.

    Suppliers sell the same part as cut tape, reel or tube, each with its own
    stock, MOQ and price ladder. Ranking has to be on the cost of the whole
    order, not on unit price: a reel often has the lower unit price but a
    5,000-piece minimum, so it is the more expensive way to obtain the 500 the
    BOM asked for. Marketplace listings are a last resort because they ship
    separately from the main order.
    """
    items = [v for v in (variations or []) if v]
    if not items:
        return None
    qty = needed if isinstance(needed, (int, float)) and needed > 0 else 1

    scored = []
    for variation in items:
        order_qty = order_quantity(qty, variation.get('minimumOrderQuantity'), variation.get('orderMultiple'))
        brk = price_at_quantity(variation.get('priceBreaks'), order_qty)
        stock = variation.get('stock')
        scored.append({
            'variation': variation,
            'orderCost': brk['unitPrice'] * order_qty if brk else None,
            'covers': stock is not None and stock >= order_qty,
        })

    priced = [s for s in scored if s['orderCost'] is not None]
    pool = priced if priced else scored
    pool.sort(key=lambda s: (
        0 if s['covers'] else 1,
        1 if s['variation'].get('marketPlace') else 0,
        s['orderCost'] if s['orderCost'] is not None else float('inf'),
        -(s['variation'].get('stock') or 0),
    ))
    return pool[0]['variation']


def build_offer(spec):
    """One supplier's answer for one BOM line, in the shape the frontend renders."""
    quantity = spec.get('quantity')
    quantity = quantity if isinstance(quantity, (int, float)) and quantity > 0 else 1
    price_breaks = spec.get('priceBreaks') or []
    moq = spec.get('minimumOrderQuantity')
    moq = moq if isinstance(moq, (int, float)) else 1
    multiple = spec.get('orderMultiple')
    multiple = multiple if isinstance(multiple, (int, float)) else 1

    order_qty = order_quantity(quantity, moq, multiple)
    brk = price_at_quantity(price_breaks, order_qty)
    unit_price = brk['unitPrice'] if brk else None
    extended_price = None if unit_price is None else round_to(unit_price * order_qty, 4)
    lead_time_days = parse_lead_time_days(spec.get('leadTime'))
    stock = spec.get('stock')
    stock = stock if isinstance(stock, (int, float)) else None
    lifecycle = normalize_lifecycle(spec.get('lifecycle'))
    lead_raw = spec.get('leadTime')
    match_count = spec.get('matchCount')

    return {
        'supplier': spec.get('supplier'),
        'found': True,
        'supplierPartNumber': spec.get('supplierPartNumber') or None,
        'manufacturer': spec.get('manufacturer') or None,
        'manufacturerPartNumber': spec.get('manufacturerPartNumber') or None,
        'description': spec.get('description') or None,
        'productUrl': spec.get('productUrl') or None,
        'datasheetUrl': spec.get('datasheetUrl') or None,
        'packaging': spec.get('packaging') or None,
        'stock': stock,
        'stockSufficient': None if stock is None else stock >= order_qty,
        'leadTimeDays': lead_time_days,
        'leadTimeText': format_lead_time(lead_time_days) or (str(lead_raw) if lead_raw else None),
        'leadTimeRaw': lead_raw if lead_raw is not None else None,
        'lifecycle': lifecycle,
        'lifecycleRaw': spec.get('lifecycle') if spec.get('lifecycle') is not None else None,
        'lifecycleSeverity': lifecycle_severity(lifecycle),
        'rohs': spec.get('rohs') or None,
        'minimumOrderQuantity': moq,
        'orderMultiple': multiple,
        'orderQuantity': order_qty,
        'unitPrice': unit_price,
        'priceBreakQuantity': brk['quantity'] if brk else None,
        'extendedPrice': extended_price,
        'currency': spec.get('currency') or 'USD',
        'priceBreaks': price_breaks,
        # How many alternates the supplier returned, so the UI can say when the
        # match was picked out of several candidates.
        'matchCount': match_count if isinstance(match_count, (int, float)) else 1,
        'exactMatch': spec.get('exactMatch') is not False,
    }


def record_to_offer(record, part):
    """Price a cached catalog record for a specific BOM line.

    A catalog record is the supplier-agnostic, quantity-independent form of one
    product. Caching that instead of a finished offer means a re-run at a
    different quantity reprices from cache without another API call.
    """
    if not record:
        return None
    variation = pick_variation(record.get('variations'), part.get('quantity'))
    variation = variation or {}
    stock = variation.get('stock')

    offer = build_offer({
        'supplier': record.get('supplier'),
        'supplierPartNumber': variation.get('supplierPartNumber') or record.get('supplierPartNumber'),
        'manufacturer': record.get('manufacturer'),
        'manufacturerPartNumber': record.get('manufacturerPartNumber'),
        'description': record.get('description'),
        'productUrl': record.get('productUrl'),
        'datasheetUrl': record.get('datasheetUrl'),
        'packaging': variation.get('packaging'),
        'stock': stock if isinstance(stock, (int, float)) else record.get('totalStock'),
        'leadTime': record.get('leadTime'),
        'lifecycle': record.get('lifecycle'),
        'rohs': record.get('rohs'),
        'quantity': part.get('quantity'),
        'minimumOrderQuantity': variation.get('minimumOrderQuantity', 1),
        'orderMultiple': variation.get('orderMultiple', 1),
        'priceBreaks': variation.get('priceBreaks') or [],
        'currency': record.get('currency'),
        'matchCount': record.get('matchCount'),
        'exactMatch': record.get('exactMatch'),
    })

    total_stock = record.get('totalStock')
    offer['totalStock'] = total_stock if isinstance(total_stock, (int, float)) else offer['stock']
    offer['packagingOptions'] = len(record.get('variations') or []) or 1
    factory_stock = record.get('factoryStock')
    if isinstance(factory_stock, (int, float)):
        offer['factoryStock'] = factory_stock
    if record.get('suggestedReplacement'):
        offer['suggestedReplacement'] = record['suggestedReplacement']

    # Pass through the extra facts an aggregator supplies that a single
    # distributor does not.
    for field in ('lifecycleRisk', 'supplyChainRisk', 'affectedByTariff'):
        if record.get(field) is not None:
            offer[field] = record[field]

    # An aggregator's variations span several distributors, so the winning one
    # names who to buy from and the rest stay available for the detail view.
    if record.get('aggregator'):
        offer['aggregator'] = True
        offer['distributor'] = (variation or {}).get('distributor')
        offer['distributorOffers'] = distributor_offers(record, part)
        offer['distributorCount'] = len(offer['distributorOffers'])
    return offer


def distributor_offers(record, part):
    """Price every distributor an aggregator returned, best first.

    Each distributor may list the part in several packagings; within one
    distributor the same total-order-cost rule picks which to quote, so the
    per-distributor rows are directly comparable with each other and with the
    single-distributor suppliers.
    """
    grouped = {}
    for variation in record.get('variations') or []:
        if not variation:
            continue
        grouped.setdefault(variation.get('distributor') or 'Unknown', []).append(variation)

    offers = []
    for name, variations in grouped.items():
        chosen = pick_variation(variations, part.get('quantity')) or {}
        stock = chosen.get('stock')
        entry = build_offer({
            'supplier': name,
            'supplierPartNumber': chosen.get('supplierPartNumber'),
            'manufacturer': record.get('manufacturer'),
            'manufacturerPartNumber': record.get('manufacturerPartNumber'),
            'description': chosen.get('description') or record.get('description'),
            'productUrl': chosen.get('productUrl') or record.get('productUrl'),
            'datasheetUrl': chosen.get('datasheetUrl') or record.get('datasheetUrl'),
            'packaging': chosen.get('packaging'),
            'stock': stock,
            'leadTime': chosen.get('leadTime'),
            'lifecycle': record.get('lifecycle'),
            'rohs': chosen.get('rohs') or record.get('rohs'),
            'quantity': part.get('quantity'),
            'minimumOrderQuantity': chosen.get('minimumOrderQuantity', 1),
            'orderMultiple': chosen.get('orderMultiple', 1),
            'priceBreaks': chosen.get('priceBreaks') or [],
            'currency': chosen.get('currency') or record.get('currency'),
            'matchCount': 1,
            'exactMatch': record.get('exactMatch'),
        })
        entry['distributor'] = name
        entry['availabilityText'] = chosen.get('availabilityText')
        entry['packagingOptions'] = len(variations)
        offers.append(entry)

    # Cheapest that can actually cover the line first; unpriced last.
    offers.sort(key=lambda o: (
        0 if o.get('stockSufficient') else 1,
        o['extendedPrice'] if o.get('extendedPrice') is not None else float('inf'),
        o.get('distributor') or '',
    ))
    return offers


def missing_offer(supplier, reason=None):
    return {
        'supplier': supplier,
        'found': False,
        'reason': reason or 'No match found',
        'stock': None,
        'leadTimeDays': None,
        'leadTimeText': None,
        'lifecycle': Lifecycle.UNKNOWN,
        'lifecycleSeverity': 'unknown',
        'unitPrice': None,
        'extendedPrice': None,
        'priceBreaks': [],
    }


def error_offer(supplier, message):
    offer = missing_offer(supplier, message)
    offer['error'] = True
    return offer


def compare_offers(offers, quantity):
    """Cross-supplier verdict for one BOM line: who wins on price, on lead time,
    on stock, and what the line's overall risk is."""
    usable = [o for o in offers if o and o.get('found')]
    summary = {
        'bestPriceSupplier': None,
        'bestPrice': None,
        'priceSpread': None,
        'priceSpreadPercent': None,
        'bestLeadTimeSupplier': None,
        'bestLeadTimeSuppliers': [],
        'bestLeadTimeDays': None,
        'inStockSuppliers': [],
        'lifecycle': worst_lifecycle([o.get('lifecycle') for o in usable]),
        'recommendedSupplier': None,
        'flags': [],
    }
    summary['lifecycleSeverity'] = lifecycle_severity(summary['lifecycle'])

    if not usable:
        summary['flags'].append({'level': 'bad', 'text': 'Not found at any configured supplier'})
        return summary

    priced = sorted(
        [o for o in usable if o.get('extendedPrice') is not None],
        key=lambda o: o['extendedPrice'],
    )
    if priced:
        summary['bestPriceSupplier'] = priced[0]['supplier']
        summary['bestPrice'] = priced[0]['extendedPrice']
        if len(priced) > 1:
            worst_price = priced[-1]['extendedPrice']
            summary['priceSpread'] = round_to(worst_price - priced[0]['extendedPrice'], 4)
            if priced[0]['extendedPrice'] > 0:
                summary['priceSpreadPercent'] = round_to(
                    ((worst_price - priced[0]['extendedPrice']) / priced[0]['extendedPrice']) * 100, 1
                )

    stocked = [o for o in usable if o.get('stockSufficient') is True]
    summary['inStockSuppliers'] = [o['supplier'] for o in stocked]

    # How soon the parts can actually arrive: a supplier holding enough stock
    # ships now regardless of what the factory quotes behind it.
    effective = []
    for offer in usable:
        if offer.get('stockSufficient') is True:
            effective.append((offer, 0))
        elif offer.get('leadTimeDays') is not None:
            effective.append((offer, offer['leadTimeDays']))

    if effective:
        fastest = min(days for _, days in effective)
        winners = [offer for offer, days in effective if days == fastest]
        summary['bestLeadTimeDays'] = fastest
        summary['bestLeadTimeSuppliers'] = [o['supplier'] for o in winners]
        # With every supplier equally fast there is nothing to single out, so
        # the UI gets no winner to badge rather than an arbitrary one.
        summary['bestLeadTimeSupplier'] = (
            winners[0]['supplier'] if len(winners) == 1 and len(winners) < len(usable) else None
        )
        priced_winners = sorted(
            [o for o in winners if o.get('extendedPrice') is not None],
            key=lambda o: o['extendedPrice'],
        )
        summary['recommendedSupplier'] = (priced_winners or winners)[0]['supplier']
    else:
        summary['recommendedSupplier'] = summary['bestPriceSupplier'] or usable[0]['supplier']

    needed = quantity if isinstance(quantity, (int, float)) and quantity > 0 else 1
    total_stock = sum(o['stock'] for o in usable if isinstance(o.get('stock'), (int, float)))

    if not stocked:
        if total_stock >= needed:
            summary['flags'].append({
                'level': 'warn',
                'text': 'No single supplier can cover %d — split the order' % needed,
            })
        else:
            summary['flags'].append({
                'level': 'bad',
                'text': 'Combined stock (%d) is below the required %d' % (total_stock, needed),
            })
    if summary['lifecycleSeverity'] == 'bad':
        summary['flags'].append({'level': 'bad', 'text': summary['lifecycle'] + ' — find a replacement'})
    elif summary['lifecycleSeverity'] == 'warn':
        summary['flags'].append({'level': 'warn', 'text': summary['lifecycle']})
    if summary['bestLeadTimeDays'] is not None and summary['bestLeadTimeDays'] >= 84:
        summary['flags'].append({
            'level': 'warn',
            'text': 'Best lead time is ' + format_lead_time(summary['bestLeadTimeDays']),
        })
    if len(usable) == 1 and len(offers) > 1:
        summary['flags'].append({
            'level': 'info',
            'text': 'Single source — only %s carries it' % usable[0]['supplier'],
        })
    if summary['priceSpreadPercent'] is not None and summary['priceSpreadPercent'] >= 15:
        summary['flags'].append({
            'level': 'info',
            'text': '%s%% price spread between suppliers' % _trim_number(summary['priceSpreadPercent']),
        })

    return summary


def _trim_number(value):
    """Render 81.1 as "81.1" and 15.0 as "15", matching the frontend's formatting."""
    if value == int(value):
        return str(int(value))
    return str(value)
