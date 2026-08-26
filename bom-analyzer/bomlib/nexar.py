"""Nexar (Altium) client, used to find alternatives to a part in trouble.

Nexar is not a supplier. The other three clients answer "what does this part
cost and can I get it"; this one answers a different question — "what could I
use instead" — and only for parts the comparison has already found to be
obsolete, end of life, not recommended, or simply unavailable. It is deliberately
never part of the BOM run: alternatives are a decision you go looking for, and
running them for every line would spend a free-tier quota on parts that are
perfectly fine.

Auth is OAuth 2.0 client credentials, the same shape DigiKey uses. The API
itself is GraphQL rather than REST, so there is one query rather than a set of
endpoints.

The query below could not be checked against a live schema while this was
written — the sandbox it was built in has no route to api.nexar.com — so it is
written from Nexar's documented Supply schema and left overridable. A wrong
field name in GraphQL fails loudly with a message naming the field, and that
message is passed straight through to the caller rather than being swallowed,
so a mismatch is a one-line fix rather than a silent empty result. Set
NEXAR_QUERY_FILE to point at your own query if the schema has moved on.
"""

import json
import os
import threading
import time
import urllib.parse

from .http_client import HttpError, request_json

SUPPLIER = 'Nexar'
TOKEN_URL = 'https://identity.nexar.com/connect/token'
API_URL = 'https://api.nexar.com/graphql'
SCOPE = 'supply.domain'

# One query, asking for the part Nexar matched and the alternatives it knows
# about. `similarParts` is Nexar's own notion of a like-for-like replacement;
# the specs come back alongside so a buyer can see *why* it is being suggested
# rather than taking the word for it.
DEFAULT_QUERY = """
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
        'the application exists but is not allowed the client-credentials grant. Check '
        'its type in the Nexar portal.',
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


class NexarClient:
    def __init__(self, client_id=None, client_secret=None, scope=None, limit=1,
                 alternatives_limit=12, token_url=None, api_url=None, query=None,
                 timeout=25.0):
        self.client_id = client_id
        self.client_secret = client_secret
        # None means "not configured", so use the default. An empty string is
        # somebody saying "ask for no scope", which is a real answer.
        self.scope = SCOPE if scope is None else (scope.strip() or None)
        self.limit = max(1, int(limit or 1))
        self.alternatives_limit = max(1, int(alternatives_limit or 12))
        self.token_url = token_url or TOKEN_URL
        self.api_url = api_url or API_URL
        self.query = query or DEFAULT_QUERY
        self.timeout = timeout
        self._token = None
        self._token_expires_at = 0.0
        self._token_lock = threading.Lock()
        # Which scope actually produced a token, once one has been minted.
        self.scope_used = None

    id = 'nexar'
    name = SUPPLIER

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

    def run_query(self, variables):
        token = self.get_token()
        result = request_json(
            self.api_url,
            method='POST',
            headers={
                'Authorization': 'Bearer ' + token,
                'Content-Type': 'application/json',
                'Accept': 'application/json',
            },
            body=json.dumps({'query': self.query, 'variables': variables}),
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


def client_from_env(env=None):
    env = env if env is not None else os.environ
    return NexarClient(
        client_id=env.get('NEXAR_CLIENT_ID'),
        client_secret=env.get('NEXAR_CLIENT_SECRET'),
        scope=env['NEXAR_SCOPE'] if 'NEXAR_SCOPE' in env else None,
        alternatives_limit=int(env.get('NEXAR_ALTERNATIVES_LIMIT') or 12),
        token_url=env.get('NEXAR_TOKEN_URL'),
        api_url=env.get('NEXAR_API_URL'),
        query=load_query(env.get('NEXAR_QUERY_FILE')),
    )
