#!/usr/bin/env python3
"""BOM Supplier Analyzer — HTTP server, static hosting and API routes.

Holds the supplier credentials and proxies the calls, because neither DigiKey
nor Mouser can be called from a browser: one needs an OAuth client secret, the
other an API key, and neither sends the CORS headers a browser requires.

Standard library only. Run with: python3 server.py
"""

import io
import json
import mimetypes
import os
import posixpath
import re
import sys
import threading
import urllib.parse
from concurrent.futures import ThreadPoolExecutor
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from bomlib.env import load_env  # noqa: E402

load_env()

from bomlib.cache import PartCache  # noqa: E402
from bomlib.digikey import DigiKeyClient  # noqa: E402
from bomlib.lookup import LookupService, summarize_bom  # noqa: E402
from bomlib.normalize import MATCH_EXACT, MATCH_MODES  # noqa: E402
from bomlib.mouser import MouserClient  # noqa: E402
from bomlib import trustedparts as tp_module  # noqa: E402
from bomlib.trustedparts import TrustedPartsClient  # noqa: E402
from bomlib.prepare import normalize_mpn, parse_prefixes, prepare_lines  # noqa: E402
from bomlib.report import write_report_workbook  # noqa: E402
from bomlib import dmsms as dmsms_module  # noqa: E402
from bomlib import nexar as nexar_module  # noqa: E402
from bomlib.spreadsheet import clean_cell, extract_bom, line_from_row, parse_workbook  # noqa: E402

PUBLIC_DIR = os.path.join(BASE_DIR, 'public')


def _int_env(name, default):
    try:
        return int(os.environ.get(name, '') or default)
    except ValueError:
        return default


def _float_env(name, default):
    try:
        return float(os.environ.get(name, '') or default)
    except ValueError:
        return default


def _bool_env(name):
    return str(os.environ.get(name, '')).strip().lower() in ('1', 'true', 'yes')


PORT = _int_env('PORT', 8787)
HOST = os.environ.get('HOST') or '0.0.0.0'
MAX_UPLOAD_BYTES = _int_env('MAX_UPLOAD_BYTES', 12 * 1024 * 1024)
MAX_JSON_BYTES = 4 * 1024 * 1024
# A finished analysis of several BOMs is far bigger than a lookup request:
# it carries every price break and distributor offer back for the report.
MAX_REPORT_BYTES = _int_env('MAX_REPORT_BYTES', 48 * 1024 * 1024)
MAX_REPORT_BOOKS = 40
MAX_PARTS_PER_REQUEST = _int_env('MAX_PARTS_PER_REQUEST', 500)
# The raw list is screened before this limit applies, so a BOM padded with
# in-house part numbers is not rejected for parts nobody would look up.
MAX_RAW_PARTS_PER_REQUEST = _int_env('MAX_RAW_PARTS_PER_REQUEST', 5000)
IGNORE_PREFIXES = parse_prefixes(os.environ.get('IGNORE_PART_PREFIXES'))


def _match_mode():
    """How closely a supplier's answer has to match the part number asked for.

    Exact by default: a keyword search will happily return a near neighbour,
    and a near neighbour priced as the real thing is worse than no answer.
    """
    mode = str(os.environ.get('MPN_MATCH') or MATCH_EXACT).strip().lower()
    return mode if mode in MATCH_MODES else MATCH_EXACT


MPN_MATCH = _match_mode()
ALLOWED_ORIGINS = [o.strip() for o in str(os.environ.get('ALLOWED_ORIGINS') or '*').split(',') if o.strip()]

digikey = DigiKeyClient(
    client_id=os.environ.get('DIGIKEY_CLIENT_ID'),
    client_secret=os.environ.get('DIGIKEY_CLIENT_SECRET'),
    sandbox=_bool_env('DIGIKEY_SANDBOX'),
    site=os.environ.get('DIGIKEY_SITE'),
    language=os.environ.get('DIGIKEY_LANGUAGE'),
    currency=os.environ.get('DIGIKEY_CURRENCY'),
    match_mode=MPN_MATCH,
)

mouser = MouserClient(
    api_key=os.environ.get('MOUSER_API_KEY'),
    currency=os.environ.get('MOUSER_CURRENCY'),
    match_mode=MPN_MATCH,
)

trustedparts = TrustedPartsClient(
    api_key=os.environ.get('TRUSTEDPARTS_API_KEY'),
    currency=os.environ.get('TRUSTEDPARTS_CURRENCY') or os.environ.get('DIGIKEY_CURRENCY'),
    country=os.environ.get('TRUSTEDPARTS_COUNTRY'),
    language=os.environ.get('TRUSTEDPARTS_LANGUAGE'),
    user_agent=os.environ.get('TRUSTEDPARTS_USER_AGENT'),
    distributors=[d.strip() for d in str(os.environ.get('TRUSTEDPARTS_DISTRIBUTORS') or '').split(',') if d.strip()],
    in_stock_only=_bool_env('TRUSTEDPARTS_IN_STOCK_ONLY'),
    use_cached_data=_bool_env('TRUSTEDPARTS_USE_CACHED_DATA'),
)

# Not a supplier: Nexar answers "what could I use instead", and only for parts
# already found to be in trouble. It is never part of a BOM run.
nexar = nexar_module.client_from_env()
MAX_ALTERNATIVE_PARTS = _int_env('MAX_ALTERNATIVE_PARTS', 50)
ALTERNATIVES_CONCURRENCY = _int_env('NEXAR_CONCURRENCY', 2)

_cache_file = os.environ.get('CACHE_FILE')
if _cache_file == 'none':
    _cache_path = None
else:
    _cache_path = _cache_file or os.path.join(BASE_DIR, '.cache', 'parts.json')

cache = PartCache(
    ttl_seconds=_float_env('CACHE_TTL_HOURS', 6) * 3600,
    path=_cache_path,
)

lookup_service = LookupService(
    clients=[digikey, mouser, trustedparts],
    cache=cache,
    concurrency=_int_env('LOOKUP_CONCURRENCY', 3),
)


def _download_name(label, suffix='-report.xlsx'):
    """A filename safe to put in a Content-Disposition header."""
    slug = re.sub(r'[^A-Za-z0-9._-]+', '-', str(label or 'bom')).strip('-') or 'bom'
    return '%s%s' % (slug[:60], suffix)


def _requested_claims(payload):
    """Part numbers an earlier BOM already claimed, mapped to its name.

    Which BOMs are open and in what order is the browser's state, not the
    server's, so the caller supplies it. Anything malformed is ignored rather
    than rejected: a bad claim map should not fail an otherwise valid lookup.
    """
    raw = payload.get('claimed')
    if not isinstance(raw, dict):
        return {}
    claims = {}
    for key, value in list(raw.items())[:MAX_RAW_PARTS_PER_REQUEST]:
        mpn = normalize_mpn(key)
        if mpn:
            claims[mpn] = str(value or 'another BOM')[:120]
    return claims


class Handler(BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.1'
    server_version = 'BOMAnalyzer/1.0'

    def log_message(self, fmt, *args):
        # The default handler logs every static asset, which drowns out the
        # lines that matter.
        pass

    # ── Routing ────────────────────────────────────────────────────────────

    def do_OPTIONS(self):
        self._apply_cors()
        self.send_response(204)
        self.send_header('Content-Length', '0')
        self.end_headers()

    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        if path.startswith('/api/'):
            self._apply_cors()
            return self._handle_api_get(path)
        return self._serve_static(path)

    def do_HEAD(self):
        self.do_GET()

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        if not path.startswith('/api/'):
            return self._send_json(405, {'error': 'Method not allowed'})
        self._apply_cors()
        if path == '/api/parse':
            return self._handle_parse()
        if path == '/api/remap':
            return self._handle_remap()
        if path == '/api/lookup':
            return self._handle_lookup()
        if path == '/api/report':
            return self._handle_report()
        if path == '/api/dmsms':
            return self._handle_dmsms()
        if path == '/api/alternatives':
            return self._handle_alternatives()
        return self._send_json(404, {'error': 'Unknown endpoint ' + path})

    def do_DELETE(self):
        path = urllib.parse.urlparse(self.path).path
        if path == '/api/cache':
            self._apply_cors()
            cache.clear()
            return self._send_json(200, {'ok': True, 'cacheEntries': len(cache)})
        self._apply_cors()
        return self._send_json(404, {'error': 'Unknown endpoint ' + path})

    def _handle_api_get(self, path):
        if path == '/api/health':
            return self._send_json(200, {
                'ok': True,
                'suppliers': [
                    {'id': 'digikey', 'name': 'DigiKey', 'configured': digikey.configured,
                     'sandbox': digikey.sandbox},
                    {'id': 'mouser', 'name': 'Mouser', 'configured': mouser.configured},
                    {'id': 'trustedparts', 'name': 'TrustedParts', 'configured': trustedparts.configured,
                     'aggregator': True,
                     # TrustedParts require visible "Powered by" attribution
                     # with a followable link wherever their data is shown.
                     'attribution': {
                         'text': tp_module.ATTRIBUTION_TEXT,
                         'name': tp_module.ATTRIBUTION_NAME,
                         'url': tp_module.ATTRIBUTION_HOME,
                         'logo': 'trustedparts-logo.svg',
                     }},
                ],
                'alternatives': {
                    'provider': nexar_module.SUPPLIER,
                    'configured': nexar.configured,
                    'maxParts': MAX_ALTERNATIVE_PARTS,
                },
                'maxPartsPerRequest': MAX_PARTS_PER_REQUEST,
                'ignorePrefixes': IGNORE_PREFIXES,
                'mpnMatch': MPN_MATCH,
                'dmsms': {
                    'statuses': list(dmsms_module.DMSMS_STATUSES),
                    'defaultSelected': list(dmsms_module.DEFAULT_SELECTED_STATUSES),
                    'resolutionOptions': list(dmsms_module.RESOLUTION_OPTIONS),
                },
                'cacheEntries': len(cache),
                'currency': os.environ.get('DIGIKEY_CURRENCY') or os.environ.get('MOUSER_CURRENCY') or 'USD',
            })
        return self._send_json(404, {'error': 'Unknown endpoint ' + path})

    # ── Endpoints ──────────────────────────────────────────────────────────

    def _handle_parse(self):
        body = self._read_body(MAX_UPLOAD_BYTES)
        if body is None:
            return self._send_json(413, {'error': 'File is too large'})
        filename = self.headers.get('X-File-Name') or 'bom.csv'
        try:
            grid = parse_workbook(body, filename)
            parsed = extract_bom(grid)
        except Exception as err:
            return self._send_json(400, {'error': 'Could not read that file: %s' % err})

        start = parsed['headerRow'] + 1
        return self._send_json(200, {
            'filename': filename,
            'headerRow': parsed['headerRow'],
            'headers': parsed['headers'],
            'mapping': parsed['mapping'],
            'lines': parsed['lines'],
            'skipped': parsed['skipped'],
            'totalRows': parsed['totalRows'],
            # The raw grid lets the UI re-derive lines when the user corrects a
            # column mapping, without a second upload. rowOffset keeps the row
            # numbers on screen pointing at the original spreadsheet rows.
            'rows': grid[start:start + 5000],
            'rowOffset': start,
        })

    def _handle_remap(self):
        payload = self._read_json()
        if payload is None:
            return
        rows = payload.get('rows') if isinstance(payload.get('rows'), list) else []
        mapping = payload.get('mapping') or {}
        mapping = {k: v for k, v in mapping.items() if isinstance(v, int)}
        offset = payload.get('rowOffset') if isinstance(payload.get('rowOffset'), int) else 0

        lines = []
        skipped = 0
        for index, row in enumerate(rows):
            if not isinstance(row, list) or all(not str(cell or '').strip() for cell in row):
                continue
            line = line_from_row(row, mapping, offset + index)
            if not line['mpn']:
                skipped += 1
                continue
            lines.append(line)
        return self._send_json(200, {'lines': lines, 'skipped': skipped})

    def _handle_lookup(self):
        payload = self._read_json()
        if payload is None:
            return
        raw_parts = payload.get('parts') if isinstance(payload.get('parts'), list) else []
        if not raw_parts:
            return self._send_json(400, {'error': 'No parts supplied'})
        if len(raw_parts) > MAX_RAW_PARTS_PER_REQUEST:
            return self._send_json(400, {
                'error': 'Send at most %d rows per request' % MAX_RAW_PARTS_PER_REQUEST
            })

        parts = []
        for i, entry in enumerate(raw_parts):
            if not isinstance(entry, dict):
                continue
            # Cleaned here as well as at parse time, so a part number typed
            # into the lookup grid or posted by any other client cannot carry
            # an invisible character into a supplier query.
            mpn = clean_cell(entry.get('mpn'))
            if not mpn:
                continue
            quantity = entry.get('quantity')
            quantity = int(quantity) if isinstance(quantity, (int, float)) and quantity > 0 else 1
            row = entry.get('row')
            parts.append({
                'row': row if isinstance(row, int) else i + 1,
                'mpn': mpn,
                'quantity': quantity,
                'manufacturer': clean_cell(entry.get('manufacturer')) or None,
                'reference': clean_cell(entry.get('reference')) or None,
                'description': clean_cell(entry.get('description')) or None,
                # The BOM's own skip-to-production column, quoted as written so
                # the reason a line was dropped can name the value.
                'skip': clean_cell(entry.get('skip')) or None,
            })

        if not parts:
            return self._send_json(400, {'error': 'No usable part numbers supplied'})

        # In-house part numbers and repeats are dropped here rather than in the
        # browser, so no caller can spend an API call on them by accident.
        #
        # A hand-entered search is the exception: somebody who types a part
        # number is asking about that part, so nothing screens it out — not the
        # in-house prefixes, not another BOM's claim, not a skip flag that could
        # only have been carried over from a BOM it did not come from.
        # Screening is there to stop automatic waste, not to refuse a direct
        # question. Repeats still merge, because typing a part twice is still
        # one part.
        manual = bool(payload.get('manual'))
        screened = prepare_lines(
            parts,
            ignore_prefixes=[] if manual else IGNORE_PREFIXES,
            claimed={} if manual else _requested_claims(payload),
            honour_skip_flag=not manual,
        )
        parts = screened['lines']
        excluded = screened['excluded']
        if not parts:
            return self._send_json(200, {
                'rows': [],
                'suppliers': lookup_service.suppliers,
                'stats': {'apiCalls': 0, 'cacheHits': 0, 'errors': 0, 'lookups': 0, 'completed': 0},
                'summary': summarize_bom([], lookup_service.suppliers),
                'excluded': excluded,
                'claimed': screened['claimed'],
            })

        # The cap applies to what will actually be looked up: a BOM padded with
        # assembly and cable lines should not be turned away for rows that were
        # never going to reach a supplier.
        if len(parts) > MAX_PARTS_PER_REQUEST:
            return self._send_json(400, {
                'error': 'Send at most %d parts per request' % MAX_PARTS_PER_REQUEST
            })

        # A large BOM takes a while, so the client can ask for server-sent
        # events and watch progress instead of staring at a spinner.
        if payload.get('stream'):
            return self._stream_lookup(parts, excluded, screened['claimed'])

        try:
            result = lookup_service.lookup_parts(parts)
        except Exception as err:
            return self._send_json(502, {'error': str(err) or 'Supplier lookup failed'})

        return self._send_json(200, {
            'rows': result['rows'],
            'suppliers': result['suppliers'],
            'stats': result['stats'],
            'summary': summarize_bom(result['rows'], result['suppliers']),
            'excluded': excluded,
            'claimed': screened['claimed'],
        })

    def _handle_report(self):
        """Build the downloadable workbook from an analysis the client already has.

        The rows come back from the browser rather than being recomputed here:
        the report has to match the numbers on screen exactly, and re-running
        the lookup would spend API calls to answer a question already answered.
        """
        payload = self._read_json(MAX_REPORT_BYTES)
        if payload is None:
            return None

        raw_books = payload.get('books')
        if not isinstance(raw_books, list) or not raw_books:
            return self._send_json(400, {'error': 'No analyzed BOMs supplied'})
        if len(raw_books) > MAX_REPORT_BOOKS:
            return self._send_json(400, {
                'error': 'Report at most %d BOMs at once' % MAX_REPORT_BOOKS
            })

        books = []
        for entry in raw_books:
            if not isinstance(entry, dict):
                continue
            rows = entry.get('rows') if isinstance(entry.get('rows'), list) else []
            suppliers = entry.get('suppliers') if isinstance(entry.get('suppliers'), list) else []
            summary = entry.get('summary') if isinstance(entry.get('summary'), dict) else {}
            if not rows:
                continue
            books.append({
                'result': {'rows': rows, 'suppliers': suppliers, 'stats': entry.get('stats') or {}},
                'summary': summary,
                'meta': {'name': str(entry.get('name') or 'Bill of materials')[:120],
                         'generated': str(entry.get('generated') or '')[:40] or None},
                'excluded': entry.get('excluded') if isinstance(entry.get('excluded'), list) else [],
            })

        if not books:
            return self._send_json(400, {'error': 'None of those BOMs have results yet'})

        buffer = io.BytesIO()
        try:
            write_report_workbook(buffer, books)
        except Exception as err:
            return self._send_json(500, {'error': 'Could not build the workbook: %s' % err})

        body = buffer.getvalue()
        name = _download_name(books[0]['meta']['name'] if len(books) == 1 else 'all-boms')
        self.send_response(200)
        self.send_header(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
        self.send_header('Content-Disposition', 'attachment; filename="%s"' % name)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        for key, value in getattr(self, '_cors', []):
            self.send_header(key, value)
        self.end_headers()
        if self.command != 'HEAD':
            self.wfile.write(body)
        return None

    def _handle_alternatives(self):
        """Ask Nexar what could be used instead of the parts named.

        Deliberately its own endpoint rather than part of /api/lookup: this runs
        only for parts somebody has decided are in trouble, and only when they
        ask for it.
        """
        if not nexar.configured:
            return self._send_json(400, {
                'error': 'Nexar is not configured. Add NEXAR_CLIENT_ID and '
                         'NEXAR_CLIENT_SECRET to .env and restart the server.',
            })

        payload = self._read_json()
        if payload is None:
            return None

        raw_parts = payload.get('parts') if isinstance(payload.get('parts'), list) else []
        parts = []
        for entry in raw_parts:
            if not isinstance(entry, dict):
                continue
            mpn = clean_cell(entry.get('mpn'))
            if not mpn:
                continue
            parts.append({'mpn': mpn, 'manufacturer': clean_cell(entry.get('manufacturer')) or None})

        if not parts:
            return self._send_json(400, {'error': 'No part numbers supplied'})
        if len(parts) > MAX_ALTERNATIVE_PARTS:
            return self._send_json(400, {
                'error': 'Ask for at most %d parts at a time' % MAX_ALTERNATIVE_PARTS,
            })

        stats = {'apiCalls': 0, 'cacheHits': 0, 'errors': 0}
        lock = threading.Lock()
        answers = {}

        def resolve(part):
            key = 'nexar %s %s' % (part['mpn'].upper(), (part['manufacturer'] or '').upper())
            cached = cache.get(key) if cache is not None else None
            if cached is not None:
                with lock:
                    stats['cacheHits'] += 1
                    answers[part['mpn']] = cached
                return
            try:
                found = nexar.find_alternatives(part)
                if cache is not None:
                    cache.set(key, found)
                with lock:
                    stats['apiCalls'] += 1
                    answers[part['mpn']] = found
            except Exception as err:
                # Not cached: a rate limit or a schema mismatch should not
                # poison the next run, and the message is what makes a schema
                # mismatch a one-line fix rather than a silent blank.
                with lock:
                    stats['apiCalls'] += 1
                    stats['errors'] += 1
                    answers[part['mpn']] = {'error': str(err) or repr(err)}

        workers = max(1, min(ALTERNATIVES_CONCURRENCY, len(parts)))
        with ThreadPoolExecutor(max_workers=workers) as pool:
            list(pool.map(resolve, parts))

        results = []
        for part in parts:
            found = answers.get(part['mpn']) or {}
            results.append({
                'mpn': part['mpn'],
                'manufacturer': part['manufacturer'],
                'matched': found.get('matched'),
                'alternatives': found.get('alternatives') or [],
                'error': found.get('error'),
            })

        return self._send_json(200, {
            'provider': nexar_module.SUPPLIER,
            'results': results,
            'stats': stats,
        })

    def _handle_dmsms(self):
        """Build a DMSMS case form for the parts the analyst ticked.

        The rows come from the browser because the selection is the whole point:
        a part can sit on three boards and belong to one program, and only the
        person filling the form knows which.
        """
        payload = self._read_json(MAX_REPORT_BYTES)
        if payload is None:
            return None

        raw_rows = payload.get('rows')
        if not isinstance(raw_rows, list) or not raw_rows:
            return self._send_json(400, {'error': 'No parts selected for the form'})

        rows = [row for row in raw_rows if isinstance(row, dict) and row.get('mpn')]
        if not rows:
            return self._send_json(400, {'error': 'None of those rows carry a part number'})

        raw_meta = payload.get('meta') if isinstance(payload.get('meta'), dict) else {}
        meta = {}
        for key in ('program', 'caseNumber', 'preparedBy', 'organization', 'contract',
                    'cage', 'date', 'notes', 'scope', 'obtained'):
            value = raw_meta.get(key)
            if value:
                meta[key] = str(value)[:400]

        buffer = io.BytesIO()
        try:
            dmsms_module.write_form(buffer, rows, meta)
        except Exception as err:
            return self._send_json(500, {'error': 'Could not build the form: %s' % err})

        body = buffer.getvalue()
        name = _download_name(meta.get('program') or 'dmsms', suffix='-dmsms.xlsx')
        self.send_response(200)
        self.send_header(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
        self.send_header('Content-Disposition', 'attachment; filename="%s"' % name)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        for key, value in getattr(self, '_cors', []):
            self.send_header(key, value)
        self.end_headers()
        if self.command != 'HEAD':
            self.wfile.write(body)
        return None

    def _stream_lookup(self, parts, excluded=None, claimed=None):
        self.send_response(200)
        self.send_header('Content-Type', 'text/event-stream; charset=utf-8')
        self.send_header('Cache-Control', 'no-cache, no-transform')
        self.send_header('X-Accel-Buffering', 'no')
        # No Content-Length is possible here, so the response is framed by the
        # connection closing.
        self.send_header('Connection', 'close')
        self.end_headers()
        self.close_connection = True

        lock = threading.Lock()
        broken = threading.Event()

        def send(event, data):
            if broken.is_set():
                return
            frame = 'event: %s\ndata: %s\n\n' % (event, json.dumps(data))
            try:
                with lock:
                    self.wfile.write(frame.encode('utf-8'))
                    self.wfile.flush()
            except (BrokenPipeError, ConnectionResetError, ValueError):
                broken.set()

        send('start', {
            'parts': len(parts),
            'suppliers': lookup_service.suppliers,
            'excluded': len(excluded or []),
        })

        # Progress fires once per supplier lookup, which is faster than the
        # client can usefully repaint; throttle to a readable rate.
        last_sent = [0.0]

        def on_progress(progress):
            import time
            now = time.time()
            if progress['completed'] == progress['total'] or now - last_sent[0] > 0.12:
                last_sent[0] = now
                send('progress', progress)

        try:
            result = lookup_service.lookup_parts(parts, on_progress=on_progress)
            send('done', {
                'rows': result['rows'],
                'suppliers': result['suppliers'],
                'stats': result['stats'],
                'summary': summarize_bom(result['rows'], result['suppliers']),
                'excluded': excluded or [],
                'claimed': claimed or [],
            })
        except Exception as err:
            send('error', {'error': str(err) or 'Supplier lookup failed'})

    # ── Static files ───────────────────────────────────────────────────────

    def _serve_static(self, path):
        relative = 'index.html' if path == '/' else path.lstrip('/')
        # Normalise away any "..", then confirm the result is still inside
        # PUBLIC_DIR before reading anything off disk.
        safe = posixpath.normpath('/' + relative).lstrip('/')
        target = os.path.abspath(os.path.join(PUBLIC_DIR, safe))
        if not (target == PUBLIC_DIR or target.startswith(PUBLIC_DIR + os.sep)):
            return self._send_json(403, {'error': 'Forbidden'})

        if not os.path.isfile(target):
            body = b'Not found'
            self.send_response(404)
            self.send_header('Content-Type', 'text/plain; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            if self.command != 'HEAD':
                self.wfile.write(body)
            return None

        content_type, _ = mimetypes.guess_type(target)
        if target.endswith('.js'):
            content_type = 'text/javascript; charset=utf-8'
        elif target.endswith('.css'):
            content_type = 'text/css; charset=utf-8'
        elif target.endswith('.html'):
            content_type = 'text/html; charset=utf-8'
        elif target.endswith('.json'):
            content_type = 'application/json; charset=utf-8'
        elif target.endswith('.svg'):
            content_type = 'image/svg+xml'

        with open(target, 'rb') as handle:
            body = handle.read()

        # The page and its script are one unit: caching them for different
        # lengths of time is how a browser ends up running last week's app.js
        # against today's index.html. "no-cache" still stores them — it just
        # revalidates first, which costs a 304 and removes the whole problem.
        # Images are versionless and genuinely static, so they keep the long TTL.
        revalidate = target.endswith(('.html', '.js', '.css', '.json'))

        self.send_response(200)
        self.send_header('Content-Type', content_type or 'application/octet-stream')
        self.send_header('Content-Length', str(len(body)))
        self.send_header(
            'Cache-Control',
            'no-cache' if revalidate else 'public, max-age=3600',
        )
        self.end_headers()
        if self.command != 'HEAD':
            self.wfile.write(body)
        return None

    # ── Plumbing ───────────────────────────────────────────────────────────

    def _apply_cors(self):
        origin = self.headers.get('Origin')
        if '*' in ALLOWED_ORIGINS:
            self._cors = [('Access-Control-Allow-Origin', '*')]
        elif origin and origin in ALLOWED_ORIGINS:
            self._cors = [('Access-Control-Allow-Origin', origin), ('Vary', 'Origin')]
        else:
            self._cors = []
        self._cors += [
            ('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS'),
            ('Access-Control-Allow-Headers', 'Content-Type, X-File-Name'),
            ('Access-Control-Max-Age', '86400'),
        ]

    def _read_body(self, limit):
        try:
            length = int(self.headers.get('Content-Length') or 0)
        except ValueError:
            return b''
        if length > limit:
            return None
        if length <= 0:
            return b''
        return self.rfile.read(length)

    def _read_json(self, limit=MAX_JSON_BYTES):
        body = self._read_body(limit)
        if body is None:
            self._send_json(413, {'error': 'Request body is too large'})
            return None
        try:
            return json.loads(body.decode('utf-8') or '{}')
        except (ValueError, UnicodeDecodeError):
            self._send_json(400, {'error': 'Request body is not valid JSON'})
            return None

    def _send_json(self, status, payload):
        body = json.dumps(payload).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        for key, value in getattr(self, '_cors', []):
            self.send_header(key, value)
        self.end_headers()
        if self.command != 'HEAD':
            self.wfile.write(body)
        return None


def build_server(host=HOST, port=PORT):
    return ThreadingHTTPServer((host, port), Handler)


def main():
    httpd = build_server()
    configured = [c.name for c in (digikey, mouser, trustedparts) if c.configured]
    print('BOM Supplier Analyzer listening on http://localhost:%d' % httpd.server_address[1])
    if configured:
        print('Suppliers configured: %s%s' % (
            ', '.join(configured),
            ' (DigiKey sandbox)' if digikey.sandbox else '',
        ))
    else:
        print('No supplier credentials found — copy .env.example to .env and add your API keys.')
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        cache.flush()
        httpd.server_close()


if __name__ == '__main__':
    main()
