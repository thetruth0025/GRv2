"""Mouser Search API v1 client.

Mouser authenticates with a single API key on the query string and reports
errors inside a 200 response, so success has to be checked in the body.
"""

import json
import re
import urllib.parse

from .http_client import HttpError, request_json
from .normalize import (
    MATCH_EXACT,
    NoMatch,
    normalize_mpn_key,
    parse_money,
    parse_quantity,
)

SUPPLIER = 'Mouser'
BASE = 'https://api.mouser.com/api/v1'


class MouserClient:
    def __init__(self, api_key=None, currency='USD', match_mode=MATCH_EXACT):
        self.api_key = api_key
        self.currency = currency or 'USD'
        self.match_mode = match_mode or MATCH_EXACT

    id = 'mouser'
    name = SUPPLIER

    @property
    def configured(self):
        return bool(self.api_key)

    def search(self, keyword, records=10):
        url = BASE + '/search/keyword?apiKey=' + urllib.parse.quote(self.api_key or '')
        result = request_json(
            url,
            method='POST',
            headers={'Content-Type': 'application/json', 'Accept': 'application/json'},
            body=json.dumps({
                'SearchByKeywordRequest': {
                    'keyword': keyword,
                    'records': records,
                    'startingRecord': 0,
                    'searchOptions': '',
                    'searchWithYourSignUpLanguage': '',
                }
            }),
        )
        data = result.get('data') or {}
        errors = data.get('Errors')
        if isinstance(errors, list) and errors:
            message = '; '.join(
                str(e.get('Message') or e.get('Code')) for e in errors if e.get('Message') or e.get('Code')
            )
            raise HttpError('Mouser API error: ' + (message or 'unknown error'), 400, errors)
        return data

    def to_record(self, data, part):
        results = (data or {}).get('SearchResults') or {}
        parts = results.get('Parts')
        if not isinstance(parts, list) or not parts:
            return NoMatch(considered=0)
        match = pick_best_part(
            parts, part.get('mpn'), part.get('manufacturer'), self.match_mode)
        if not match:
            near = nearest_part(parts, part.get('mpn'))
            return NoMatch(
                closest=(near or {}).get('ManufacturerPartNumber'),
                manufacturer=(near or {}).get('Manufacturer'),
                considered=len(parts),
            )
        return build_record(match, part, self.currency, len(parts))

    def fetch_record(self, part):
        return self.to_record(self.search(part.get('mpn')), part)


normalize_key = normalize_mpn_key


def stock_of(part):
    """Availability arrives as prose ("1,234 In Stock", "None");
    AvailabilityInStock is the clean number when Mouser includes it."""
    direct = parse_quantity(part.get('AvailabilityInStock'))
    if direct is not None:
        return direct
    availability = str(part.get('Availability') or '').strip()
    if not availability or re.fullmatch(r'none', availability, re.IGNORECASE):
        return 0
    parsed = parse_quantity(availability)
    return parsed if parsed is not None else 0


def pick_best_part(parts, mpn, manufacturer=None, mode=MATCH_EXACT):
    """Choose the part that is the requested part, or say none was.

    Mouser's search widens to related parts, so in exact mode only the part
    number itself will do. Manufacturer and stock break ties between several
    spellings of the same number; they never promote a different number.
    """
    want_mpn = normalize_key(mpn)
    if not want_mpn:
        return None
    want_mfr = normalize_key(manufacturer)

    best = None
    best_score = -1
    for candidate in parts:
        candidate_mpn = normalize_key(candidate.get('ManufacturerPartNumber'))
        candidate_mfr = normalize_key(candidate.get('Manufacturer'))
        score = 0
        if candidate_mpn and candidate_mpn == want_mpn:
            score += 100
        elif mode == MATCH_EXACT:
            continue
        elif candidate_mpn and (candidate_mpn.startswith(want_mpn) or want_mpn.startswith(candidate_mpn)):
            score += 50
        elif candidate_mpn and (want_mpn in candidate_mpn or candidate_mpn in want_mpn):
            score += 20
        if want_mfr and candidate_mfr and (
            candidate_mfr == want_mfr or want_mfr in candidate_mfr or candidate_mfr in want_mfr
        ):
            score += 30
        if stock_of(candidate) > 0:
            score += 5
        if score > best_score:
            best_score = score
            best = candidate

    floor = 100 if mode == MATCH_EXACT else 20
    return best if best_score >= floor else None


def nearest_part(parts, mpn):
    """The returned part number closest to the request, for the message only."""
    want = normalize_key(mpn)
    best = None
    best_shared = -1
    for candidate in parts:
        key = normalize_key(candidate.get('ManufacturerPartNumber'))
        if not key:
            continue
        shared = 0
        for a, b in zip(key, want):
            if a != b:
                break
            shared += 1
        if shared > best_shared:
            best_shared = shared
            best = candidate
    return best


def price_breaks_of(part):
    items = part.get('PriceBreaks')
    if not isinstance(items, list):
        return []
    breaks = []
    for entry in items:
        quantity = parse_quantity(entry.get('Quantity'))
        unit_price = parse_money(entry.get('Price'))
        if quantity is None or unit_price is None:
            continue
        breaks.append({'quantity': quantity, 'unitPrice': unit_price, 'currency': entry.get('Currency')})
    return sorted(breaks, key=lambda b: b['quantity'])


def lifecycle_of(part):
    """LifecycleStatus is authoritative but frequently blank; ProductStatus is
    the Mouser-catalog view ("New at Mouser", "Obsolete") and fills the gap."""
    lifecycle = str(part.get('LifecycleStatus') or '').strip()
    if lifecycle:
        return lifecycle
    status = str(part.get('ProductStatus') or '').strip()
    return status or None


def build_record(part, bom_part, currency, match_count):
    breaks = price_breaks_of(part)
    stock = stock_of(part)
    record = {
        'supplier': SUPPLIER,
        'manufacturer': part.get('Manufacturer'),
        'manufacturerPartNumber': part.get('ManufacturerPartNumber'),
        'description': part.get('Description'),
        'productUrl': part.get('ProductDetailUrl'),
        'datasheetUrl': part.get('DataSheetUrl'),
        'leadTime': part.get('LeadTime'),
        'lifecycle': lifecycle_of(part),
        'rohs': part.get('ROHSStatus'),
        'totalStock': stock,
        'currency': (breaks[0].get('currency') if breaks else None) or currency,
        'matchCount': match_count,
        'exactMatch': normalize_key(part.get('ManufacturerPartNumber')) == normalize_key(bom_part.get('mpn')),
        'factoryStock': parse_quantity(part.get('FactoryStock')),
        # Mouser sells one catalog line per part number, so there is a single
        # packaging option rather than DigiKey's variation list.
        'variations': [{
            'supplierPartNumber': part.get('MouserPartNumber'),
            'packaging': 'Reel' if part.get('Reeling') else None,
            'stock': stock,
            'minimumOrderQuantity': parse_quantity(part.get('Min')) or 1,
            'orderMultiple': parse_quantity(part.get('Mult')) or 1,
            'priceBreaks': breaks,
            'marketPlace': False,
        }],
    }
    replacement = str(part.get('SuggestedReplacement') or '').strip()
    if replacement:
        record['suggestedReplacement'] = replacement
    return record
