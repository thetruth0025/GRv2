"""TrustedParts.com Inventory API v2 client.

Unlike DigiKey and Mouser, TrustedParts is an aggregator: one part number comes
back with a list of authorized distributors, each with its own stock, pricing
and packaging. Every distributor is kept — the comparison table quotes the best
one and the detail view lists them all.

Two things the API does not provide, which the rest of the app has to tolerate:

* **No lead time.** A distributor result carries stock only, so lead time is
  reported as unknown rather than guessed at. Lines that TrustedParts stocks
  still compare correctly, because stock on hand outranks a quoted lead time.
* **Lifecycle is a risk rating, not a status,** and is gated by TrustedParts'
  own policy (`LifecycleRisk` is null unless your account is approved for it).
  It is surfaced verbatim as a risk field and only feeds the lifecycle column
  when the text actually matches lifecycle vocabulary — a "High" risk rating is
  not the same claim as "Obsolete", so it is never rendered as one.

Spec: POST https://api.trustedparts.com/v2/search, key in the X-Api-Key header.
"""

import json
import re

from .http_client import HttpError, request_json
from .normalize import (
    Lifecycle,
    NoMatch,
    normalize_lifecycle,
    parse_money,
    parse_quantity,
)

SUPPLIER = 'TrustedParts'
BASE = 'https://api.trustedparts.com'
SITE = 'https://www.trustedparts.com'

# TrustedParts require that a publicly available application displaying their
# data shows "Powered by" followed by their logo, linked back to them, and that
# the link is followable — no rel="nofollow", no temporary redirect.
ATTRIBUTION_TEXT = 'Powered by'
ATTRIBUTION_NAME = 'TrustedParts.com'
ATTRIBUTION_HOME = SITE

# The API accepts up to 50 search queries per request.
MAX_QUERIES_PER_REQUEST = 50


class TrustedPartsClient:
    def __init__(self, api_key=None, currency='USD', country='US', language='en',
                 user_agent=None, distributors=None, in_stock_only=False, use_cached_data=False):
        self.api_key = api_key
        self.currency = currency or 'USD'
        self.country = country or 'US'
        self.language = language or 'en'
        self.user_agent = user_agent or 'BOM-Supplier-Analyzer/1.0'
        # Optional allow-list; omitted entirely means "all authorized distributors".
        self.distributors = [d for d in (distributors or []) if d]
        self.in_stock_only = bool(in_stock_only)
        self.use_cached_data = bool(use_cached_data)

    id = 'trustedparts'
    name = SUPPLIER
    # Signals to LookupService that this client prefers batched lookups.
    batch_size = MAX_QUERIES_PER_REQUEST

    @property
    def configured(self):
        return bool(self.api_key)

    def search(self, parts):
        """One request covering up to `batch_size` parts."""
        queries = []
        for part in parts:
            token = str(part.get('mpn') or '').strip()
            if not token:
                continue
            # The API rejects tokens outside 2-100 characters.
            token = token[:100]
            if len(token) < 2:
                continue
            query = {'SearchToken': token}
            manufacturer = str(part.get('manufacturer') or '').strip()
            if manufacturer:
                query['Manufacturers'] = [manufacturer]
            queries.append(query)

        if not queries:
            return {}

        body = {
            'Queries': queries,
            'CurrencyCode': self.currency,
            'CountryCode': self.country,
            'LanguageCode': self.language,
            # A BOM lookup wants the part asked for, not near misses. The API
            # forces exact matching on multi-part requests anyway.
            'ExactMatch': True,
            'InStockOnly': self.in_stock_only,
            'UseCachedData': self.use_cached_data,
            'IsCrawler': False,
            'UserAgent': self.user_agent,
        }
        if self.distributors:
            body['Distributors'] = self.distributors

        result = request_json(
            BASE + '/v2/search',
            method='POST',
            headers={
                'X-Api-Key': self.api_key,
                'Content-Type': 'application/json',
                'Accept': 'application/json',
            },
            body=json.dumps(body),
        )
        data = result.get('data') or {}

        # Errors arrive in the body rather than as a non-2xx status.
        error = str(data.get('ErrorMessage') or '').strip()
        if error:
            raise HttpError('TrustedParts API error: ' + error, 400, data)
        return data

    def to_records(self, data, parts):
        """Map one response onto the requested parts, keyed by normalized MPN."""
        results = (data or {}).get('PartResults')
        if not isinstance(results, list):
            return {}

        links = collect_attribution_links(data)
        by_key = {}
        for result in results:
            if not isinstance(result, dict):
                continue
            key = normalize_key(result.get('PartNumber'))
            if not key:
                continue
            # A search token can match more than one manufacturer's part; keep
            # the first and let manufacturer scoring below choose.
            by_key.setdefault(key, []).append(result)

        records = {}
        for part in parts:
            want = normalize_key(part.get('mpn'))
            candidates = by_key.get(want) or []
            best = pick_best_result(candidates, part.get('manufacturer'))
            if best is None:
                # Already exact — this lookup is keyed on the normalized part
                # number, so a miss means the number was not among the results.
                # Say which numbers were, so a near miss is not read as absent.
                records[part.get('mpn')] = NoMatch(
                    closest=nearest_result(results, part.get('mpn')),
                    considered=len(results),
                )
                continue
            record = build_record(best, part, self.currency, links)
            if record:
                records[part.get('mpn')] = record
        return records

    def fetch_records(self, parts):
        """Batch entry point used by LookupService."""
        return self.to_records(self.search(parts), parts)

    def fetch_record(self, part):
        """Single-part fallback, so the client also satisfies the plain protocol."""
        return self.fetch_records([part]).get(part.get('mpn'))


def normalize_key(text):
    return re.sub(r'[^A-Z0-9]', '', str(text or '').upper())


def absolute_url(url):
    """Attribution links come back as site-relative paths."""
    text = str(url or '').strip()
    if not text:
        return None
    if text.startswith('http://') or text.startswith('https://'):
        return text
    return SITE + ('' if text.startswith('/') else '/') + text


def collect_attribution_links(payload):
    """Pull the attribution Links section out of a response.

    The published OpenAPI schema does not describe this section, but the
    attribution documentation does, so it is read defensively: from the
    response root or from a part result, tolerating both the documented
    {Key, SearchToken, Manufacturer, Url} shape and the {Type, Url} shape the
    schema uses elsewhere. A response without it simply yields no links.
    """
    primary = None
    by_token = {}

    def absorb(links):
        nonlocal primary
        if not isinstance(links, list):
            return
        for link in links:
            if not isinstance(link, dict):
                continue
            url = absolute_url(link.get('Url'))
            if not url:
                continue
            key = str(link.get('Key') or '').strip()
            if key.lower() == 'primary':
                primary = primary or url
                continue
            token = normalize_key(link.get('SearchToken'))
            if token:
                by_token.setdefault(token, url)

    if isinstance(payload, dict):
        absorb(payload.get('Links'))
        for result in payload.get('PartResults') or []:
            if isinstance(result, dict):
                absorb(result.get('Links'))

    return {'primary': primary, 'byToken': by_token}


def nearest_result(results, mpn):
    """The returned part number closest to the request, for the message only."""
    want = normalize_key(mpn)
    best = None
    best_shared = -1
    for result in results or []:
        number = (result or {}).get('PartNumber')
        key = normalize_key(number)
        if not key:
            continue
        shared = 0
        for a, b in zip(key, want):
            if a != b:
                break
            shared += 1
        if shared > best_shared:
            best_shared = shared
            best = number
    return best


def pick_best_result(results, manufacturer=None):
    """Prefer the requested manufacturer, then the result carrying most stock."""
    if not results:
        return None
    want = normalize_key(manufacturer)

    def score(result):
        mfr = normalize_key(result.get('Manufacturer'))
        matched = bool(want and mfr and (mfr == want or want in mfr or mfr in want))
        return (0 if matched else 1, -total_stock_of(result))

    return sorted(results, key=score)[0]


def total_stock_of(result):
    total = 0
    for distributor in result.get('Distributors') or []:
        for entry in (distributor or {}).get('DistributorResults') or []:
            quantity = stock_of(entry)
            if quantity:
                total += quantity
    return total


def stock_of(entry):
    """QuantityOnHand is null when TrustedParts withholds the exact figure; the
    Availability prose is the only signal left in that case."""
    stock = (entry or {}).get('Stock') or {}
    quantity = stock.get('QuantityOnHand')
    if isinstance(quantity, (int, float)):
        return int(quantity)
    availability = str(stock.get('Availability') or '').strip()
    if not availability:
        return None
    if re.search(r'\bno(ne|t)?\b|out of stock|unavailable', availability, re.IGNORECASE):
        return 0
    parsed = parse_quantity(availability)
    # "In Stock" with no number means stocked at an undisclosed depth, which is
    # not the same as zero, so it stays unknown rather than becoming a number.
    return parsed if parsed is not None else None


def price_breaks_of(entry):
    pricing = (entry or {}).get('Pricing') or {}
    breaks = []
    for price in pricing.get('Prices') or []:
        if not isinstance(price, dict):
            continue
        quantity = parse_quantity(price.get('Quantity'))
        amount = price.get('Amount')
        unit_price = amount if isinstance(amount, (int, float)) else parse_money(
            price.get('FormattedAmount') or price.get('Text')
        )
        if quantity is None or unit_price is None:
            continue
        breaks.append({'quantity': quantity, 'unitPrice': float(unit_price)})
    return sorted(breaks, key=lambda b: b['quantity'])


def link_of(entry, *types):
    for link in (entry or {}).get('Links') or []:
        if not isinstance(link, dict):
            continue
        kind = str(link.get('Type') or '').strip().lower()
        if kind in types and link.get('Url'):
            return link['Url']
    return None


def rohs_of(entry):
    compliance = (entry or {}).get('Compliance') or {}
    for item in compliance.get('RoHS') or []:
        if not isinstance(item, dict):
            continue
        description = str(item.get('Description') or '').strip()
        if description:
            return description
        if item.get('IsCompliant') is not None:
            return 'RoHS Compliant' if item['IsCompliant'] else 'Not RoHS Compliant'
    return None


def packaging_of(entry):
    """Packaging is a list; take the first named type and any minimum it sets."""
    for package in (entry or {}).get('Packaging') or []:
        if not isinstance(package, dict):
            continue
        name = str(package.get('PackageType') or '').strip()
        moq = parse_quantity(package.get('MinimumOrderQuantity'))
        if name or moq:
            return name or None, moq
    return None, None


def build_record(result, part, currency, links=None):
    """Flatten PartResults → Distributors → DistributorResults into variations.

    Each variation carries the distributor that offers it, so the shared
    packaging-selection logic ranks across every distributor at once and the
    per-distributor breakdown can be regrouped from the same list.
    """
    variations = []
    for distributor in result.get('Distributors') or []:
        if not isinstance(distributor, dict):
            continue
        name = str(distributor.get('Name') or '').strip() or 'Unknown distributor'
        for entry in distributor.get('DistributorResults') or []:
            if not isinstance(entry, dict):
                continue
            pricing = entry.get('Pricing') or {}
            package_name, package_moq = packaging_of(entry)
            moq = parse_quantity(pricing.get('MinimumQuantity')) or package_moq or 1
            multiple = parse_quantity(pricing.get('QuantityMultiple')) or 1
            stock = stock_of(entry)
            variations.append({
                'distributor': name,
                'distributorId': distributor.get('Id'),
                'supplierPartNumber': entry.get('DistributorPartNumber') or None,
                'packaging': package_name,
                'stock': stock,
                'availabilityText': ((entry.get('Stock') or {}).get('Availability') or None),
                'minimumOrderQuantity': moq,
                'orderMultiple': multiple,
                'priceBreaks': price_breaks_of(entry),
                'currency': pricing.get('CurrencyCode') or currency,
                'description': entry.get('Description') or None,
                'datasheetUrl': link_of(entry, 'datasheet'),
                'productUrl': link_of(entry, 'distributor', 'product'),
                'rohs': rohs_of(entry),
                'marketPlace': False,
            })

    if not variations:
        return None

    known_stock = [v['stock'] for v in variations if isinstance(v['stock'], (int, float))]
    lifecycle_risk = str(result.get('LifecycleRisk') or '').strip() or None

    # Attribution target for this line: the part's own TrustedParts page when
    # the response names one, else the product page, else their home page.
    links = links or {'primary': None, 'byToken': {}}
    part_url = (
        links['byToken'].get(normalize_key(result.get('PartNumber')))
        or links['byToken'].get(normalize_key(part.get('mpn')))
        or absolute_url(result.get('ProductUrl'))
        or links.get('primary')
        or ATTRIBUTION_HOME
    )

    return {
        'supplier': SUPPLIER,
        'aggregator': True,
        'manufacturer': result.get('Manufacturer') or None,
        'manufacturerPartNumber': result.get('PartNumber') or None,
        'description': next((v['description'] for v in variations if v['description']), None),
        # The TrustedParts page for this part, which is also the attribution
        # link back to them for this row.
        'productUrl': part_url,
        'attribution': {
            'text': ATTRIBUTION_TEXT,
            'name': ATTRIBUTION_NAME,
            'url': part_url,
            'home': ATTRIBUTION_HOME,
        },
        'datasheetUrl': next((v['datasheetUrl'] for v in variations if v['datasheetUrl']), None),
        # The API reports no lead time, so none is claimed.
        'leadTime': None,
        'lifecycle': lifecycle_status_from_risk(lifecycle_risk),
        'lifecycleRisk': lifecycle_risk,
        'supplyChainRisk': str(result.get('SupplyChainRisk') or '').strip() or None,
        'affectedByTariff': bool(result.get('IsAffectedByTariff')),
        'rohs': next((v['rohs'] for v in variations if v['rohs']), None),
        'totalStock': sum(known_stock) if known_stock else None,
        'currency': currency,
        'matchCount': len(variations),
        'exactMatch': normalize_key(result.get('PartNumber')) == normalize_key(part.get('mpn')),
        'variations': variations,
    }


def lifecycle_status_from_risk(raw):
    """Only promote LifecycleRisk to a lifecycle status when it says one.

    TrustedParts reports a risk rating, which may be a grade like "Low" rather
    than a status. Mapping a grade onto "Obsolete" would assert something the
    API never said, so anything unrecognized stays Unknown.
    """
    if not raw:
        return None
    status = normalize_lifecycle(raw)
    return None if status == Lifecycle.UNKNOWN else raw
