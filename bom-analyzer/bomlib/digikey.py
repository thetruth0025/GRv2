"""DigiKey Product Information V4 client.

DigiKey uses OAuth 2.0 client credentials. Tokens are short-lived (10 minutes
in practice), so one is held in memory and refreshed just before it expires
rather than fetched per part.
"""

import json
import re
import threading
import time
import urllib.parse

from .http_client import HttpError, request_json
from .normalize import (
    MATCH_EXACT,
    NoMatch,
    mpn_equal,
    normalize_mpn_key,
    parse_money,
    parse_quantity,
)

SUPPLIER = 'DigiKey'
PROD_BASE = 'https://api.digikey.com'
SANDBOX_BASE = 'https://sandbox-api.digikey.com'


class DigiKeyClient:
    def __init__(self, client_id=None, client_secret=None, sandbox=False,
                 site='US', language='en', currency='USD', match_mode=MATCH_EXACT):
        self.client_id = client_id
        self.client_secret = client_secret
        self.sandbox = bool(sandbox)
        self.base_url = SANDBOX_BASE if sandbox else PROD_BASE
        self.site = site or 'US'
        self.language = language or 'en'
        self.currency = currency or 'USD'
        self.match_mode = match_mode or MATCH_EXACT
        self._token = None
        self._token_expires_at = 0.0
        self._token_lock = threading.Lock()

    id = 'digikey'
    name = SUPPLIER

    @property
    def configured(self):
        return bool(self.client_id and self.client_secret)

    def get_token(self):
        # Collapse concurrent refreshes so a burst of lookups mints one token.
        with self._token_lock:
            if self._token and time.time() < self._token_expires_at:
                return self._token

            body = urllib.parse.urlencode({
                'client_id': self.client_id,
                'client_secret': self.client_secret,
                'grant_type': 'client_credentials',
            })
            result = request_json(
                self.base_url + '/v1/oauth2/token',
                method='POST',
                headers={'Content-Type': 'application/x-www-form-urlencoded'},
                body=body,
                retries=1,
            )
            data = result.get('data') or {}
            if not data.get('access_token'):
                raise HttpError('DigiKey token response did not contain an access_token', 0, data)

            self._token = data['access_token']
            try:
                lifetime = float(data.get('expires_in') or 600)
            except (TypeError, ValueError):
                lifetime = 600.0
            self._token_expires_at = time.time() + max(30.0, lifetime - 60.0)
            return self._token

    def search(self, keyword, limit=10):
        token = self.get_token()
        result = request_json(
            self.base_url + '/products/v4/search/keyword',
            method='POST',
            headers={
                'Authorization': 'Bearer ' + token,
                'X-DIGIKEY-Client-Id': self.client_id,
                'X-DIGIKEY-Locale-Site': self.site,
                'X-DIGIKEY-Locale-Language': self.language,
                'X-DIGIKEY-Locale-Currency': self.currency,
                'Content-Type': 'application/json',
                'Accept': 'application/json',
            },
            body=json.dumps({'Keywords': keyword, 'Limit': limit, 'Offset': 0}),
        )
        return result.get('data') or {}

    def to_record(self, data, part):
        """A catalog record, or a NoMatch naming what came back instead."""
        products = collect_products(data)
        if not products:
            return NoMatch(considered=0)
        product = pick_best_product(
            products, part.get('mpn'), part.get('manufacturer'), self.match_mode)
        if not product:
            near = nearest_product(products, part.get('mpn'))
            return NoMatch(
                closest=mpn_of(near) if near else None,
                manufacturer=manufacturer_of(near) if near else None,
                considered=len(products),
            )
        return build_record(product, part, self.currency, len(products))

    def fetch_record(self, part):
        return self.to_record(self.search(part.get('mpn')), part)


def collect_products(data):
    out = []
    seen = set()

    def push(items):
        if not isinstance(items, list):
            return
        for item in items:
            if not isinstance(item, dict):
                continue
            key = (mpn_of(item), manufacturer_of(item))
            if key in seen:
                continue
            seen.add(key)
            out.append(item)

    # ExactMatches first: V4 puts high-confidence hits there.
    push((data or {}).get('ExactMatches'))
    push((data or {}).get('Products'))
    push((data or {}).get('ExactManufacturerProducts'))
    return out


def mpn_of(product):
    return str(product.get('ManufacturerProductNumber') or product.get('ManufacturerPartNumber') or '')


def manufacturer_of(product):
    mfr = product.get('Manufacturer')
    if not mfr:
        return ''
    if isinstance(mfr, str):
        return mfr
    return str(mfr.get('Name') or mfr.get('Value') or '')


def description_of(product):
    desc = product.get('Description')
    if isinstance(desc, str):
        return desc
    if isinstance(desc, dict):
        return desc.get('ProductDescription') or desc.get('DetailedDescription')
    return product.get('ProductDescription') or product.get('DetailedDescription')


def status_of(product):
    status = product.get('ProductStatus')
    if isinstance(status, str) and status.strip():
        return status
    if isinstance(status, dict):
        value = status.get('Status') or status.get('Value')
        if value:
            return value
    # Fall back to the boolean flags V4 also exposes.
    if product.get('Obsolete'):
        return 'Obsolete'
    if product.get('EndOfLife'):
        return 'End of Life'
    if product.get('Discontinued'):
        return 'Discontinued'
    if product.get('NormallyStocking'):
        return 'Active'
    return None


# Kept as a module name because it is what this file has always called it.
normalize_key = normalize_mpn_key


def pick_best_product(products, keyword, manufacturer=None, mode=MATCH_EXACT):
    """Choose the product that is the requested part, or say none was.

    A keyword search returns loosely related parts — ask for LM358 and LM358DR
    comes back, which is a different device in a different package. In exact
    mode only the part number itself will do; the manufacturer and stock break
    ties between several spellings of the same number, and never promote a
    different number.
    """
    want_mpn = normalize_key(keyword)
    if not want_mpn:
        return None
    want_mfr = normalize_key(manufacturer)

    best = None
    best_score = -1
    for product in products:
        mpn = normalize_key(mpn_of(product))
        mfr = normalize_key(manufacturer_of(product))
        score = 0
        if mpn and mpn == want_mpn:
            score += 100
        elif mode == MATCH_EXACT:
            # Anything that is not the number asked for is not the part.
            continue
        elif mpn and (mpn.startswith(want_mpn) or want_mpn.startswith(mpn)):
            score += 50
        elif mpn and (want_mpn in mpn or mpn in want_mpn):
            score += 20
        if want_mfr and mfr and (mfr == want_mfr or want_mfr in mfr or mfr in want_mfr):
            score += 30
        if (parse_quantity(product.get('QuantityAvailable')) or 0) > 0:
            score += 5
        if score > best_score:
            best_score = score
            best = product

    floor = 100 if mode == MATCH_EXACT else 20
    return best if best_score >= floor else None


def nearest_product(products, keyword):
    """The returned part number closest to what was asked for, for the message.

    Only used to explain a rejection, so "closest" here means longest shared
    opening — enough to tell a missing part from one spelled differently.
    """
    want = normalize_key(keyword)
    best = None
    best_shared = -1
    for product in products:
        mpn = mpn_of(product)
        key = normalize_key(mpn)
        if not key:
            continue
        shared = 0
        for a, b in zip(key, want):
            if a != b:
                break
            shared += 1
        if shared > best_shared:
            best_shared = shared
            best = product
    return best


def price_breaks_of(items):
    if not isinstance(items, list):
        return []
    breaks = []
    for entry in items:
        quantity = parse_quantity(entry.get('BreakQuantity'))
        unit_price = parse_money(entry.get('UnitPrice'))
        if quantity is None or unit_price is None:
            continue
        breaks.append({'quantity': quantity, 'unitPrice': unit_price})
    return sorted(breaks, key=lambda b: b['quantity'])


def variations_of(product):
    """V4 nests packaging options under ProductVariations, each with its own
    DigiKey part number, stock, MOQ and price ladder. V3 flattened all of that
    onto the product, so both shapes are handled."""
    variations = product.get('ProductVariations')
    if isinstance(variations, list) and variations:
        out = []
        for v in variations:
            package = v.get('PackageType') or {}
            stock = parse_quantity(v.get('QuantityAvailableforPackageType'))
            if stock is None:
                stock = parse_quantity(v.get('QuantityAvailable'))
            out.append({
                'supplierPartNumber': v.get('DigiKeyProductNumber') or v.get('DigiKeyPartNumber'),
                'packaging': (package.get('Name') or package.get('Value')) if isinstance(package, dict) else None,
                'stock': stock,
                'minimumOrderQuantity': parse_quantity(v.get('MinimumOrderQuantity')) or 1,
                'orderMultiple': parse_quantity(v.get('StandardPackage')) or 1,
                'priceBreaks': price_breaks_of(v.get('StandardPricing')),
                'marketPlace': bool(v.get('MarketPlace')),
            })
        return out

    packaging = product.get('Packaging') or {}
    return [{
        'supplierPartNumber': product.get('DigiKeyPartNumber') or product.get('DigiKeyProductNumber'),
        'packaging': (packaging.get('Name') or packaging.get('Value')) if isinstance(packaging, dict) else None,
        'stock': parse_quantity(product.get('QuantityAvailable')),
        'minimumOrderQuantity': parse_quantity(product.get('MinimumOrderQuantity')) or 1,
        'orderMultiple': parse_quantity(product.get('StandardPackage')) or 1,
        'priceBreaks': price_breaks_of(product.get('StandardPricing')),
        'marketPlace': False,
    }]


def build_record(product, part, currency, match_count):
    classifications = product.get('Classifications') or {}
    return {
        'supplier': SUPPLIER,
        'manufacturer': manufacturer_of(product) or None,
        'manufacturerPartNumber': mpn_of(product) or None,
        'description': description_of(product),
        'productUrl': product.get('ProductUrl'),
        'datasheetUrl': product.get('DatasheetUrl'),
        'leadTime': product.get('ManufacturerLeadWeeks') or product.get('LeadStatus'),
        'lifecycle': status_of(product),
        'rohs': classifications.get('RohsStatus'),
        'totalStock': parse_quantity(product.get('QuantityAvailable')),
        'currency': currency,
        'matchCount': match_count,
        'exactMatch': normalize_key(mpn_of(product)) == normalize_key(part.get('mpn')),
        'variations': variations_of(product),
    }
