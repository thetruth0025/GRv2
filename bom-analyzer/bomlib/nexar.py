"""Nexar (Altium) client: a supplier column, and a source of alternatives.

Nexar answers two questions, and this client asks both.

**"Can I get this part, and what does it cost."** Like TrustedParts, Nexar is an
aggregator rather than a distributor: one part number comes back with a list of
sellers, each with its own stock, price ladder and packaging. It takes its place
alongside DigiKey, Mouser and TrustedParts as another column in the comparison,
built from the same catalog-record shape, so nothing downstream needs to know
where the numbers came from.

**"What could I use instead."** Nexar's `similarParts` is asked separately, on
demand, and only for parts the comparison has already found to be in trouble.
That is a decision you go looking for, so it is never part of a BOM run.

The two are independent. `similarParts` is not on every Nexar plan, and a plan
that refuses it still answers part search perfectly well — which is the point of
keeping them apart rather than behind one query.

Auth is OAuth 2.0 client credentials, the same shape DigiKey uses. The API
itself is GraphQL rather than REST, so there are queries rather than endpoints.

Neither query could be checked against a live schema while this was written —
the sandbox it was built in has no route to api.nexar.com — so both are written
from Nexar's documented Supply schema and left overridable. A wrong field name
in GraphQL fails loudly with a message naming the field, and that message is
passed straight through rather than being swallowed, so a mismatch is a one-line
fix and never a silent empty result. Part search additionally falls back on its
own: if the batched `supMultiMatch` query is rejected by the schema, the client
drops to per-part `supSearchMpn` for the rest of the run rather than reporting
every line as not carried. Set NEXAR_SEARCH_QUERY_FILE or NEXAR_QUERY_FILE to
point at your own query if the schema has moved on.
"""

import json
import os
import re
import threading
import time
import urllib.parse

from .http_client import HttpError, request_json
from .normalize import (
    MATCH_EXACT,
    NoMatch,
    normalize_lifecycle,
    normalize_mpn_key,
    Lifecycle,
)

SUPPLIER = 'Nexar'
TOKEN_URL = 'https://identity.nexar.com/connect/token'
API_URL = 'https://api.nexar.com/graphql'
SCOPE = 'supply.domain'

# Nexar's Supply data is Octopart's, and their terms ask that it be credited.
ATTRIBUTION_TEXT = 'Part data via'
ATTRIBUTION_NAME = 'Nexar (Octopart)'
ATTRIBUTION_HOME = 'https://octopart.com'

# supMultiMatch takes many queries in one request, which is what a BOM is. The
# free tier is metered, so asking once for fifty parts rather than fifty times
# for one is the difference between a quota that lasts a month and one that
# does not.
MAX_QUERIES_PER_REQUEST = 20

# The fields a supplier column needs: who sells it, how many they hold, what
# they charge at each break, and what you have to buy to get one.
PART_FIELDS = """
    mpn
    manufacturer { name }
    shortDescription
    octopartUrl
    bestDatasheet { url }
    totalAvail
    estimatedFactoryLeadDays
    medianPrice1000 { price currency }
    specs { attribute { name shortname } displayValue }
    sellers {
      company { name }
      isAuthorized
      offers {
        sku
        inventoryLevel
        moq
        orderMultiple
        packaging
        clickUrl
        factoryLeadDays
        prices { quantity price currency convertedPrice convertedCurrency }
      }
    }
"""

# Batched part search: one request, one entry per BOM line. `reference` comes
# back on each hit, which is what pairs an answer with the line that asked.
SEARCH_QUERY = """
query BomMultiMatch($queries: [SupPartMatchQuery!]!) {
  supMultiMatch(queries: $queries) {
    reference
    hits
    parts {
%s
    }
  }
}
""" % PART_FIELDS

# The fallback, used per part when the batched query is not in the schema.
SEARCH_ONE_QUERY = """
query BomPartSearch($q: String!, $limit: Int!) {
  supSearchMpn(q: $q, limit: $limit) {
    results {
      part {
%s
      }
    }
  }
}
""" % PART_FIELDS


# One query, asking for the part Nexar matched and the alternatives it knows
# about. `similarParts` is Nexar's own notion of a like-for-like replacement;
# the specs come back alongside so a buyer can see *why* it is being suggested
# rather than taking the word for it. Not on every Nexar plan — a plan without
# it still answers part search, which is why the two are separate queries.
ALTERNATIVES_QUERY = """
query BomAlternatives($q: String!, $limit: Int!) {
  supSearchMpn(q: $q, limit: $limit) {
    results {
      part {
        mpn
        manufacturer { name }
        shortDescription
        octopartUrl
        totalAvail
        estimatedFactoryLeadDays
        medianPrice1000 { price currency }
        bestDatasheet { url }
        specs { attribute { name shortname } displayValue }
        similarParts {
          mpn
          manufacturer { name }
          shortDescription
          octopartUrl
          totalAvail
          estimatedFactoryLeadDays
          medianPrice1000 { price currency }
          bestDatasheet { url }
          specs { attribute { name shortname } displayValue }
        }
      }
    }
  }
}
"""


# The name this query went by when alternatives were all Nexar did here.
DEFAULT_QUERY = ALTERNATIVES_QUERY


# What an OAuth error code from the token endpoint actually means for someone
# holding a .env file. RFC 6749 names the codes; none of them say what to do.
TOKEN_HINTS = {
    'invalid_client':
        'NEXAR_CLIENT_ID or NEXAR_CLIENT_SECRET is wrong. Copy both again from the '
        'application page — a secret truncated on paste looks exactly like this.',
    'invalid_scope':
        'the application is not granted the scope being asked for. Set NEXAR_SCOPE to a '
        'scope it does have, or leave NEXAR_SCOPE empty to ask for none.',
    'unauthorized_client':
        'the credentials are real, but this application is not allowed to authenticate as '
        'itself. A Nexar Design application signs a user in instead, which is a different '
        'grant and a different API. Alternatives come from the Supply API, so what is '
        'needed is an application with Supply access — add it to this one in the Nexar '
        'portal, or create a second application for it.',
    'unsupported_grant_type':
        'the token endpoint did not accept client_credentials. Check NEXAR_TOKEN_URL.',
    'invalid_request':
        'the token request was malformed — if NEXAR_TOKEN_URL is set, check it points at '
        "Nexar's /connect/token endpoint.",
}


def oauth_error_code(error):
    """The `error` field of an OAuth failure body, if there is one."""
    body = getattr(error, 'body', None)
    if isinstance(body, dict):
        return str(body.get('error') or '').strip()
    return ''


def describe_token_failure(error, scope):
    """Turn "HTTP 400 from identity.nexar.com" into something actionable.

    The token endpoint answers a failure with a body naming the reason. Raising
    only the status throws that away and leaves nothing to act on, which is the
    one thing this wrapper exists to prevent.
    """
    body = getattr(error, 'body', None)
    body = body if isinstance(body, dict) else {}
    code = str(body.get('error') or '').strip()
    detail = str(body.get('error_description') or '').strip()

    pieces = ['Nexar refused the credentials']
    if code:
        pieces.append('(%s)' % code)
    if scope:
        pieces.append('while asking for scope "%s"' % scope)
    message = ' '.join(pieces) + '.'

    if detail:
        message += ' ' + detail.rstrip('.') + '.'
    hint = TOKEN_HINTS.get(code)
    if hint:
        message += ' ' + hint[0].upper() + hint[1:]
    elif not code:
        # No parseable body: say what came back so it is not a bare status.
        raw = body or getattr(error, 'body', None)
        message += ' The endpoint returned: %s' % (raw if raw else str(error))
    return message


def _text(value):
    return str(value).strip() if value not in (None, '') else None


def _number(value):
    if isinstance(value, bool) or value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _integer(value):
    number = _number(value)
    return int(number) if number is not None else None


def part_from_node(node):
    """One Nexar part node, flattened into the shape the app displays.

    Every field is read defensively: the schema carries a lot of optional
    relations, and a part with no datasheet or no price is ordinary, not an
    error.
    """
    if not isinstance(node, dict):
        return None

    mpn = _text(node.get('mpn'))
    if not mpn:
        return None

    price = node.get('medianPrice1000') or {}
    datasheet = node.get('bestDatasheet') or {}
    manufacturer = node.get('manufacturer') or {}

    specs = []
    for spec in node.get('specs') or []:
        if not isinstance(spec, dict):
            continue
        attribute = spec.get('attribute') or {}
        name = _text(attribute.get('shortname')) or _text(attribute.get('name'))
        value = _text(spec.get('displayValue'))
        if name and value:
            specs.append({'name': name, 'value': value})

    return {
        'mpn': mpn,
        'manufacturer': _text(manufacturer.get('name')),
        'description': _text(node.get('shortDescription')),
        'url': _text(node.get('octopartUrl')),
        'datasheetUrl': _text(datasheet.get('url')),
        'stock': _integer(node.get('totalAvail')),
        'leadDays': _integer(node.get('estimatedFactoryLeadDays')),
        'medianPrice': _number(price.get('price')),
        'currency': _text(price.get('currency')),
        'specs': specs,
    }


# ── Part search: turning a Nexar part into a catalog record ─────────────────

normalize_key = normalize_mpn_key

# What Nexar calls the field. It arrives as a spec rather than a column, so it
# is looked for by name rather than assumed to be in a fixed place.
LIFECYCLE_SPEC_NAMES = ('lifecyclestatus', 'lifecycle_status', 'lifecycle status',
                        'lifecycle', 'partstatus', 'part status')


def mpn_of(node):
    return _text((node or {}).get('mpn'))


def manufacturer_of(node):
    return _text(((node or {}).get('manufacturer') or {}).get('name'))


def specs_of(node):
    """Nexar specs, flattened to name/value pairs."""
    out = []
    for spec in (node or {}).get('specs') or []:
        if not isinstance(spec, dict):
            continue
        attribute = spec.get('attribute') or {}
        name = _text(attribute.get('shortname')) or _text(attribute.get('name'))
        value = _text(spec.get('displayValue'))
        if name and value:
            out.append({'name': name, 'value': value})
    return out


def lifecycle_of(node):
    """The lifecycle status, if Nexar reported one among the specs.

    Anything it does not recognise as a status stays unset rather than being
    rendered as one: an unfamiliar spec value is not a claim about supply.
    """
    for spec in specs_of(node):
        if str(spec['name']).replace(' ', '').lower().rstrip('s') not in (
                n.replace(' ', '').replace('_', '').rstrip('s') for n in LIFECYCLE_SPEC_NAMES):
            continue
        if normalize_lifecycle(spec['value']) != Lifecycle.UNKNOWN:
            return spec['value']
    return None


def price_breaks_of(offer, fallback_currency):
    """One offer's price ladder, cheapest-per-piece ordering left to the caller.

    Nexar returns both the seller's own currency and a converted figure. The
    converted one is used when it exists, because a column that mixes
    currencies is not a comparison.
    """
    breaks = []
    for entry in (offer or {}).get('prices') or []:
        if not isinstance(entry, dict):
            continue
        quantity = _integer(entry.get('quantity'))
        price = _number(entry.get('convertedPrice'))
        currency = _text(entry.get('convertedCurrency'))
        if price is None:
            price = _number(entry.get('price'))
            currency = _text(entry.get('currency'))
        if quantity is None or quantity < 1 or price is None:
            continue
        breaks.append({
            'quantity': quantity,
            'unitPrice': price,
            'currency': currency or fallback_currency,
        })
    breaks.sort(key=lambda b: b['quantity'])
    return breaks


def variations_of(node, currency, authorized_only=False):
    """Flatten sellers → offers into the packaging options the app ranks.

    Each offer is one seller's way of selling the part — a reel, a cut tape, a
    tube — so it maps onto the same variation shape TrustedParts produces and
    is ranked by the same total-order-cost rule.
    """
    variations = []
    for seller in (node or {}).get('sellers') or []:
        if not isinstance(seller, dict):
            continue
        authorized = seller.get('isAuthorized')
        if authorized_only and authorized is False:
            continue
        name = _text((seller.get('company') or {}).get('name')) or 'Unknown seller'
        for offer in seller.get('offers') or []:
            if not isinstance(offer, dict):
                continue
            breaks = price_breaks_of(offer, currency)
            variations.append({
                'distributor': name,
                'supplierPartNumber': _text(offer.get('sku')),
                'packaging': _text(offer.get('packaging')),
                'stock': _integer(offer.get('inventoryLevel')),
                'minimumOrderQuantity': _integer(offer.get('moq')) or 1,
                'orderMultiple': _integer(offer.get('orderMultiple')) or 1,
                'priceBreaks': breaks,
                'currency': (breaks[0]['currency'] if breaks else currency),
                'productUrl': _text(offer.get('clickUrl')),
                'leadDays': _integer(offer.get('factoryLeadDays')),
                # Nexar flags an unauthorized seller; the ranking treats those
                # the way it treats a marketplace listing — a last resort.
                'marketPlace': authorized is False,
            })
    return variations


def lead_time_of(node, variations):
    """Days to the factory, said in the words the lead-time parser expects."""
    days = _integer((node or {}).get('estimatedFactoryLeadDays'))
    if days is None:
        offered = [v['leadDays'] for v in variations if isinstance(v.get('leadDays'), int)]
        days = min(offered) if offered else None
    return None if days is None else '%d Days' % days


def build_record(node, part, currency, match_count=1):
    """One Nexar part, in the supplier-agnostic shape the app prices later."""
    variations = variations_of(node, currency)
    if not variations:
        # Nexar knows the part but nobody listed is selling it. That is a real
        # answer — "not carried" — rather than a match.
        return None

    known_stock = [v['stock'] for v in variations if isinstance(v['stock'], int)]
    total = _integer((node or {}).get('totalAvail'))
    part_url = _text((node or {}).get('octopartUrl'))

    return {
        'supplier': SUPPLIER,
        'aggregator': True,
        'manufacturer': manufacturer_of(node),
        'manufacturerPartNumber': mpn_of(node),
        'description': _text((node or {}).get('shortDescription')),
        'productUrl': part_url,
        'attribution': {
            'text': ATTRIBUTION_TEXT,
            'name': ATTRIBUTION_NAME,
            'url': part_url or ATTRIBUTION_HOME,
            'home': ATTRIBUTION_HOME,
        },
        'datasheetUrl': _text(((node or {}).get('bestDatasheet') or {}).get('url')),
        'leadTime': lead_time_of(node, variations),
        'lifecycle': lifecycle_of(node),
        'totalStock': total if total is not None else (sum(known_stock) if known_stock else None),
        'currency': currency,
        'matchCount': match_count,
        'exactMatch': normalize_key(mpn_of(node)) == normalize_key(part.get('mpn')),
        'variations': variations,
    }


def pick_best_part(nodes, keyword, manufacturer=None, mode=MATCH_EXACT):
    """The node that is the part asked for, or None.

    Same rule as the other clients: in exact mode nothing but the part number
    itself will do, and the manufacturer only breaks ties between spellings of
    that number — it never promotes a different one.
    """
    want_mpn = normalize_key(keyword)
    if not want_mpn:
        return None
    want_mfr = normalize_key(manufacturer)

    best = None
    best_score = -1
    for node in nodes:
        mpn = normalize_key(mpn_of(node))
        mfr = normalize_key(manufacturer_of(node))
        score = 0
        if mpn and mpn == want_mpn:
            score += 100
        elif mode == MATCH_EXACT:
            continue
        elif mpn and (mpn.startswith(want_mpn) or want_mpn.startswith(mpn)):
            score += 50
        elif mpn and (want_mpn in mpn or mpn in want_mpn):
            score += 20
        if want_mfr and mfr and (mfr == want_mfr or want_mfr in mfr or mfr in want_mfr):
            score += 30
        if (_integer(node.get('totalAvail')) or 0) > 0:
            score += 5
        if score > best_score:
            best_score = score
            best = node

    floor = 100 if mode == MATCH_EXACT else 20
    return best if best_score >= floor else None


def nearest_part(nodes, keyword):
    """The returned part number closest to the one asked for, for the message."""
    want = normalize_key(keyword)
    best = None
    best_shared = -1
    for node in nodes:
        key = normalize_key(mpn_of(node))
        if not key:
            continue
        shared = 0
        for a, b in zip(key, want):
            if a != b:
                break
            shared += 1
        if shared > best_shared:
            best_shared = shared
            best = node
    return best


# A GraphQL rejection that means "this query is not in the schema" rather than
# "your input was wrong". Nexar words it as an unknown field or unknown type;
# either way the query can never succeed, so retrying it is waste.
_SCHEMA_MISS = re.compile(
    r"cannot query field|unknown (?:field|type|argument)|"
    r"is not defined by type|no field named|not available on your plan",
    re.I,
)


def _query_not_in_schema(error):
    return bool(_SCHEMA_MISS.search(str(error or '')))


class NexarClient:
    def __init__(self, client_id=None, client_secret=None, scope=None, limit=1,
                 alternatives_limit=12, token_url=None, api_url=None, query=None,
                 search_query=None, currency='USD', match_mode=MATCH_EXACT,
                 batch_size=MAX_QUERIES_PER_REQUEST, timeout=25.0):
        self.client_id = client_id
        self.client_secret = client_secret
        # None means "not configured", so use the default. An empty string is
        # somebody saying "ask for no scope", which is a real answer.
        self.scope = SCOPE if scope is None else (scope.strip() or None)
        self.limit = max(1, int(limit or 1))
        self.alternatives_limit = max(1, int(alternatives_limit or 12))
        self.token_url = token_url or TOKEN_URL
        self.api_url = api_url or API_URL
        self.query = query or ALTERNATIVES_QUERY
        self.search_query = search_query or SEARCH_QUERY
        self.currency = currency or 'USD'
        self.match_mode = match_mode or MATCH_EXACT
        self.batch_size = max(1, int(batch_size or MAX_QUERIES_PER_REQUEST))
        self.timeout = timeout
        self._token = None
        self._token_expires_at = 0.0
        self._token_lock = threading.Lock()
        # Which scope actually produced a token, once one has been minted.
        self.scope_used = None
        # Set once the batched query is found not to be in the schema, so the
        # rest of the run goes straight to the per-part fallback.
        self._batch_unsupported = False
        # GraphQL requests actually sent, so a run that fell back to per-part
        # queries reports the quota it really spent rather than one per batch.
        self.requests_made = 0

    id = 'nexar'
    name = SUPPLIER
    aggregator = True

    @property
    def configured(self):
        return bool(self.client_id and self.client_secret)

    def request_token(self, scope):
        """One token exchange. Omits the scope entirely when there is none."""
        fields = {
            'grant_type': 'client_credentials',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
        }
        if scope:
            fields['scope'] = scope

        try:
            result = request_json(
                self.token_url,
                method='POST',
                headers={'Content-Type': 'application/x-www-form-urlencoded'},
                body=urllib.parse.urlencode(fields),
                timeout=self.timeout,
                retries=1,
            )
        except HttpError as err:
            raise HttpError(describe_token_failure(err, scope), err.status, err.body)
        return result.get('data') or {}

    def get_token(self):
        # Collapse concurrent refreshes so a burst of lookups mints one token,
        # exactly as the DigiKey client does.
        with self._token_lock:
            if self._token and time.time() < self._token_expires_at:
                return self._token

            try:
                data = self.request_token(self.scope)
                self.scope_used = self.scope
            except HttpError as err:
                # Much the commonest 400 here: the application was never granted
                # the scope being asked for. Nexar issues a usable token without
                # one, so that is worth trying before giving up on the run.
                if not (self.scope and oauth_error_code(err) == 'invalid_scope'):
                    raise
                try:
                    data = self.request_token(None)
                except HttpError as second:
                    # Report the refusal that names the scope, since that is the
                    # actionable one, and say the fallback was tried as well.
                    raise HttpError(
                        '%s Requesting a token with no scope failed too (%s).'
                        % (err, oauth_error_code(second) or 'no reason given'),
                        second.status, second.body,
                    )
                self.scope_used = None

            if not data.get('access_token'):
                raise HttpError('Nexar token response did not contain an access_token', 0, data)

            self._token = data['access_token']
            try:
                lifetime = float(data.get('expires_in') or 3600)
            except (TypeError, ValueError):
                lifetime = 3600.0
            self._token_expires_at = time.time() + max(30.0, lifetime - 60.0)
            return self._token

    def run_query(self, variables, query=None):
        token = self.get_token()
        self.requests_made += 1
        result = request_json(
            self.api_url,
            method='POST',
            headers={
                'Authorization': 'Bearer ' + token,
                'Content-Type': 'application/json',
                'Accept': 'application/json',
            },
            body=json.dumps({'query': query or self.query, 'variables': variables}),
            timeout=self.timeout,
            retries=1,
        )
        data = result.get('data')
        if not isinstance(data, dict):
            raise HttpError('Nexar returned a response that was not JSON', 0, data)

        # GraphQL answers 200 with an errors array rather than an HTTP status,
        # so a schema mismatch would otherwise look like "no alternatives".
        errors = data.get('errors')
        if errors:
            messages = []
            for error in errors if isinstance(errors, list) else [errors]:
                if isinstance(error, dict) and error.get('message'):
                    messages.append(str(error['message']))
            raise HttpError(
                'Nexar rejected the query: %s' % ('; '.join(messages) or 'no message given'),
                0, errors,
            )
        return data.get('data') or {}

    # ── Part search ─────────────────────────────────────────────────────────

    def match_query(self, part, reference):
        """One entry of a supMultiMatch request.

        The manufacturer is sent as its own field rather than glued onto the
        part number: a jellybean number is made by several people, and Nexar
        can only narrow on it if it is told which field it is.
        """
        query = {'mpn': str(part.get('mpn') or '').strip(),
                 'reference': reference,
                 'limit': self.limit}
        manufacturer = str(part.get('manufacturer') or '').strip()
        if manufacturer:
            query['manufacturer'] = manufacturer
        return query

    def search(self, parts):
        """One batched request covering up to `batch_size` parts.

        Returns {reference: [part nodes]}. A line Nexar had nothing for is
        absent rather than present and empty, which `to_records` reads as a
        miss.
        """
        queries = []
        wanted = []
        for index, part in enumerate(parts):
            if not str(part.get('mpn') or '').strip():
                continue
            reference = 'line-%d' % index
            queries.append(self.match_query(part, reference))
            wanted.append(reference)
        if not queries:
            return {}

        payload = self.run_query({'queries': queries}, query=self.search_query)
        found = {}
        for hit in payload.get('supMultiMatch') or []:
            if not isinstance(hit, dict):
                continue
            reference = _text(hit.get('reference'))
            nodes = [n for n in (hit.get('parts') or []) if isinstance(n, dict)]
            if reference:
                found[reference] = nodes
        return found

    def search_one(self, part):
        """The per-part fallback, used when supMultiMatch is not in the schema."""
        term = self.search_term(part)
        if not term:
            return []
        payload = self.run_query({'q': term, 'limit': self.limit}, query=SEARCH_ONE_QUERY)
        results = ((payload.get('supSearchMpn') or {}).get('results')) or []
        nodes = []
        for entry in results:
            node = (entry or {}).get('part') if isinstance(entry, dict) else None
            if isinstance(node, dict):
                nodes.append(node)
        return nodes

    def to_record(self, nodes, part):
        """A catalog record, or a NoMatch naming what came back instead."""
        nodes = [n for n in (nodes or []) if isinstance(n, dict)]
        if not nodes:
            return NoMatch(considered=0)
        node = pick_best_part(nodes, part.get('mpn'), part.get('manufacturer'), self.match_mode)
        if not node:
            near = nearest_part(nodes, part.get('mpn'))
            return NoMatch(
                closest=mpn_of(near) if near else None,
                manufacturer=manufacturer_of(near) if near else None,
                considered=len(nodes),
            )
        record = build_record(node, part, self.currency, len(nodes))
        if record is None:
            # Nexar knows the part; nobody it lists is selling it.
            return NoMatch(considered=len(nodes))
        return record

    def fetch_records(self, parts):
        """Batch entry point used by LookupService, keyed by the MPN asked for."""
        parts = list(parts)
        if not self._batch_unsupported:
            try:
                found = self.search(parts)
            except HttpError as err:
                if not _query_not_in_schema(err):
                    raise
                # The plan or the schema does not have supMultiMatch. Reporting
                # every line as not carried would be a lie about availability,
                # so drop to the query that does exist and stay there.
                self._batch_unsupported = True
            else:
                return {part.get('mpn'): self.to_record(found.get('line-%d' % index), part)
                        for index, part in enumerate(parts)}
        return {part.get('mpn'): self.fetch_record(part) for part in parts}

    def fetch_record(self, part):
        """Single-part entry point, on the query every plan has."""
        return self.to_record(self.search_one(part), part)

    # ── Alternatives ────────────────────────────────────────────────────────

    def search_term(self, part):
        """What to ask Nexar for. The manufacturer narrows a part number that
        several makers use, which is common for jellybean parts."""
        pieces = [str(part.get('mpn') or '').strip()]
        manufacturer = str(part.get('manufacturer') or '').strip()
        if manufacturer:
            pieces.append(manufacturer)
        return ' '.join(piece for piece in pieces if piece)

    def find_alternatives(self, part):
        """Return {'matched': part-or-None, 'alternatives': [...]}.

        The matched part comes back alongside the alternatives on purpose: a
        suggestion is only worth as much as the match it came from, and a buyer
        needs to see that Nexar found the right part before trusting what it
        offers instead.
        """
        term = self.search_term(part)
        if not term:
            return {'matched': None, 'alternatives': []}

        payload = self.run_query({'q': term, 'limit': self.limit})
        results = ((payload.get('supSearchMpn') or {}).get('results')) or []

        matched = None
        raw = None
        for entry in results:
            node = (entry or {}).get('part') if isinstance(entry, dict) else None
            candidate = part_from_node(node)
            if candidate:
                matched, raw = candidate, node
                break

        if matched is None:
            return {'matched': None, 'alternatives': []}

        alternatives = []
        seen = {matched['mpn'].upper()}
        for node in (raw or {}).get('similarParts') or []:
            alternative = part_from_node(node)
            if not alternative:
                continue
            key = alternative['mpn'].upper()
            if key in seen:
                continue
            seen.add(key)
            alternatives.append(alternative)
            if len(alternatives) >= self.alternatives_limit:
                break

        return {'matched': matched, 'alternatives': alternatives}


def load_query(path):
    """Read a query override, so a schema change does not need a code change."""
    if not path:
        return None
    with open(path, 'r', encoding='utf-8') as handle:
        text = handle.read().strip()
    return text or None


def client_from_env(env=None, match_mode=MATCH_EXACT):
    env = env if env is not None else os.environ
    try:
        batch = int(env.get('NEXAR_BATCH_SIZE') or MAX_QUERIES_PER_REQUEST)
    except ValueError:
        batch = MAX_QUERIES_PER_REQUEST
    return NexarClient(
        client_id=env.get('NEXAR_CLIENT_ID'),
        client_secret=env.get('NEXAR_CLIENT_SECRET'),
        scope=env['NEXAR_SCOPE'] if 'NEXAR_SCOPE' in env else None,
        alternatives_limit=int(env.get('NEXAR_ALTERNATIVES_LIMIT') or 12),
        token_url=env.get('NEXAR_TOKEN_URL'),
        api_url=env.get('NEXAR_API_URL'),
        query=load_query(env.get('NEXAR_QUERY_FILE')),
        search_query=load_query(env.get('NEXAR_SEARCH_QUERY_FILE')),
        currency=env.get('NEXAR_CURRENCY') or env.get('DIGIKEY_CURRENCY') or 'USD',
        match_mode=match_mode,
        batch_size=batch,
    )
