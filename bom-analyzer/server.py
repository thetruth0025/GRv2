#!/usr/bin/env python3
"""BOM Supplier Analyzer — HTTP server, static hosting and API routes.

Holds the supplier credentials and proxies the calls, because neither DigiKey
nor Mouser can be called from a browser: one needs an OAuth client secret, the
other an API key, and neither sends the CORS headers a browser requires.

Standard library only. Run with: python3 server.py
"""

import json
import mimetypes
import os
import posixpath
import sys
import threading
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from bomlib.env import load_env  # noqa: E402

load_env()

from bomlib.cache import PartCache  # noqa: E402
from bomlib.digikey import DigiKeyClient  # noqa: E402
from bomlib.lookup import LookupService, summarize_bom  # noqa: E402
from bomlib.mouser import MouserClient  # noqa: E402
from bomlib.trustedparts import TrustedPartsClient  # noqa: E402
from bomlib.spreadsheet import extract_bom, line_from_row, parse_workbook  # noqa: E402

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
MAX_PARTS_PER_REQUEST = _int_env('MAX_PARTS_PER_REQUEST', 500)
ALLOWED_ORIGINS = [o.strip() for o in str(os.environ.get('ALLOWED_ORIGINS') or '*').split(',') if o.strip()]

digikey = DigiKeyClient(
    client_id=os.environ.get('DIGIKEY_CLIENT_ID'),
    client_secret=os.environ.get('DIGIKEY_CLIENT_SECRET'),
    sandbox=_bool_env('DIGIKEY_SANDBOX'),
    site=os.environ.get('DIGIKEY_SITE'),
    language=os.environ.get('DIGIKEY_LANGUAGE'),
    currency=os.environ.get('DIGIKEY_CURRENCY'),
)

mouser = MouserClient(
    api_key=os.environ.get('MOUSER_API_KEY'),
    currency=os.environ.get('MOUSER_CURRENCY'),
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
                     'aggregator': True},
                ],
                'maxPartsPerRequest': MAX_PARTS_PER_REQUEST,
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
        if len(raw_parts) > MAX_PARTS_PER_REQUEST:
            return self._send_json(400, {
                'error': 'Send at most %d parts per request' % MAX_PARTS_PER_REQUEST
            })

        parts = []
        for i, entry in enumerate(raw_parts):
            if not isinstance(entry, dict):
                continue
            mpn = str(entry.get('mpn') or '').strip()
            if not mpn:
                continue
            quantity = entry.get('quantity')
            quantity = int(quantity) if isinstance(quantity, (int, float)) and quantity > 0 else 1
            row = entry.get('row')
            parts.append({
                'row': row if isinstance(row, int) else i + 1,
                'mpn': mpn,
                'quantity': quantity,
                'manufacturer': str(entry['manufacturer']).strip() if entry.get('manufacturer') else None,
                'reference': str(entry['reference']).strip() if entry.get('reference') else None,
                'description': str(entry['description']).strip() if entry.get('description') else None,
            })

        if not parts:
            return self._send_json(400, {'error': 'No usable part numbers supplied'})

        # A large BOM takes a while, so the client can ask for server-sent
        # events and watch progress instead of staring at a spinner.
        if payload.get('stream'):
            return self._stream_lookup(parts)

        try:
            result = lookup_service.lookup_parts(parts)
        except Exception as err:
            return self._send_json(502, {'error': str(err) or 'Supplier lookup failed'})

        return self._send_json(200, {
            'rows': result['rows'],
            'suppliers': result['suppliers'],
            'stats': result['stats'],
            'summary': summarize_bom(result['rows'], result['suppliers']),
        })

    def _stream_lookup(self, parts):
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

        send('start', {'parts': len(parts), 'suppliers': lookup_service.suppliers})

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

        self.send_response(200)
        self.send_header('Content-Type', content_type or 'application/octet-stream')
        self.send_header('Content-Length', str(len(body)))
        self.send_header(
            'Cache-Control',
            'no-cache' if target.endswith('.html') else 'public, max-age=3600',
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

    def _read_json(self):
        body = self._read_body(MAX_JSON_BYTES)
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
