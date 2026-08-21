import json
import os
import threading
import unittest
import urllib.error
import urllib.request

# Keep the test run away from the developer's real credentials and cache file.
os.environ['CACHE_FILE'] = 'none'
for name in ('DIGIKEY_CLIENT_ID', 'DIGIKEY_CLIENT_SECRET', 'MOUSER_API_KEY'):
    os.environ.pop(name, None)

import server  # noqa: E402


def call(path, method='GET', body=None, headers=None):
    url = BASE + path
    data = body.encode('utf-8') if isinstance(body, str) else body
    request = urllib.request.Request(url, data=data, method=method)
    for key, value in (headers or {}).items():
        request.add_header(key, value)
    try:
        with urllib.request.urlopen(request, timeout=10) as response:
            text = response.read().decode('utf-8')
            status = response.status
            content_type = response.headers.get('Content-Type')
    except urllib.error.HTTPError as err:
        text = err.read().decode('utf-8')
        status = err.code
        content_type = err.headers.get('Content-Type')
    try:
        parsed = json.loads(text) if text else None
    except ValueError:
        parsed = None
    return {'status': status, 'data': parsed, 'text': text, 'contentType': content_type}


BASE = None
_httpd = None
_thread = None


def setUpModule():
    global BASE, _httpd, _thread
    _httpd = server.build_server('127.0.0.1', 0)
    _thread = threading.Thread(target=_httpd.serve_forever, daemon=True)
    _thread.start()
    BASE = 'http://127.0.0.1:%d' % _httpd.server_address[1]


def tearDownModule():
    _httpd.shutdown()
    _httpd.server_close()


class HealthTests(unittest.TestCase):
    def test_reports_which_suppliers_are_configured(self):
        result = call('/api/health')
        self.assertEqual(result['status'], 200)
        self.assertTrue(result['data']['ok'])
        self.assertEqual([s['id'] for s in result['data']['suppliers']], ['digikey', 'mouser'])
        # No credentials are set in this environment.
        self.assertTrue(all(s['configured'] is False for s in result['data']['suppliers']))


class ParseTests(unittest.TestCase):
    def test_uploading_a_csv_returns_columns_and_lines(self):
        csv = '\n'.join([
            'Item,Reference,Qty,Manufacturer,Manufacturer Part Number,Description',
            '1,C1,300,Murata,GRM188R71H104KA93D,CAP CER 0.1UF',
            '2,R1,500,Yageo,RC0603FR-0710KL,RES 10K',
        ])
        result = call('/api/parse', 'POST', csv, {
            'X-File-Name': 'bom.csv', 'Content-Type': 'application/octet-stream',
        })
        self.assertEqual(result['status'], 200)
        self.assertEqual(len(result['data']['lines']), 2)
        self.assertEqual(result['data']['lines'][0]['mpn'], 'GRM188R71H104KA93D')
        self.assertEqual(result['data']['mapping']['mpn'], 4)
        self.assertIsInstance(result['data']['rows'], list)

    def test_remapping_re_derives_lines_without_a_second_upload(self):
        rows = [
            ['1', 'C1', '300', 'Murata', 'GRM188R71H104KA93D'],
            ['2', 'R1', '500', 'Yageo', 'RC0603FR-0710KL'],
        ]
        result = call('/api/remap', 'POST', json.dumps({
            'rows': rows, 'mapping': {'mpn': 4, 'quantity': 2, 'reference': 1}, 'rowOffset': 1,
        }), {'Content-Type': 'application/json'})
        self.assertEqual(result['status'], 200)
        self.assertEqual(len(result['data']['lines']), 2)
        self.assertEqual(result['data']['lines'][1]['mpn'], 'RC0603FR-0710KL')
        self.assertEqual(result['data']['lines'][1]['quantity'], 500)
        self.assertEqual(result['data']['lines'][1]['reference'], 'R1')

    def test_remapped_rows_keep_original_row_numbers(self):
        csv = '\n'.join([
            'Acme Widget rev C',
            '',
            'Item,Reference,Qty,Manufacturer Part Number',
            '1,C1,300,GRM188R71H104KA93D',
            '2,R1,500,RC0603FR-0710KL',
        ])
        parsed = call('/api/parse', 'POST', csv, {
            'X-File-Name': 'bom.csv', 'Content-Type': 'application/octet-stream',
        })
        # The header is on spreadsheet row 3, so the first part is on row 4.
        self.assertEqual(parsed['data']['lines'][0]['row'], 4)

        remapped = call('/api/remap', 'POST', json.dumps({
            'rows': parsed['data']['rows'],
            'mapping': parsed['data']['mapping'],
            'rowOffset': parsed['data']['rowOffset'],
        }), {'Content-Type': 'application/json'})
        self.assertEqual(
            [l['row'] for l in remapped['data']['lines']],
            [l['row'] for l in parsed['data']['lines']],
        )

    def test_unreadable_upload_is_rejected_with_a_readable_message(self):
        result = call('/api/parse', 'POST', b'PK\x03\x04\x00\x00', {
            'X-File-Name': 'broken.xlsx', 'Content-Type': 'application/octet-stream',
        })
        self.assertEqual(result['status'], 400)
        self.assertIn('Could not read that file', result['data']['error'])


class LookupTests(unittest.TestCase):
    def test_lookup_with_no_parts_is_a_client_error(self):
        result = call('/api/lookup', 'POST', json.dumps({'parts': []}),
                      {'Content-Type': 'application/json'})
        self.assertEqual(result['status'], 400)
        self.assertIn('No parts supplied', result['data']['error'])

    def test_lookup_with_no_configured_supplier_explains_what_is_missing(self):
        result = call('/api/lookup', 'POST',
                      json.dumps({'parts': [{'mpn': 'RC0603FR-0710KL', 'quantity': 10}]}),
                      {'Content-Type': 'application/json'})
        self.assertEqual(result['status'], 502)
        self.assertIn('No supplier is configured', result['data']['error'])

    def test_malformed_json_is_rejected_without_taking_the_server_down(self):
        result = call('/api/lookup', 'POST', '{not json', {'Content-Type': 'application/json'})
        self.assertEqual(result['status'], 400)
        self.assertIn('not valid JSON', result['data']['error'])


class StaticTests(unittest.TestCase):
    def test_frontend_is_served_from_the_same_origin_as_the_api(self):
        result = call('/')
        self.assertEqual(result['status'], 200)
        self.assertIn('BOM Supplier Analyzer', result['text'])

    def test_path_traversal_outside_public_is_refused(self):
        result = call('/../server.py')
        self.assertIn(result['status'], (403, 404))
        self.assertNotIn('DIGIKEY_CLIENT_SECRET', result['text'])

    def test_unknown_api_route_returns_404_not_the_frontend(self):
        result = call('/api/nope')
        self.assertEqual(result['status'], 404)
        self.assertIn('Unknown endpoint', result['data']['error'])


class StreamingTests(unittest.TestCase):
    """The streaming path is how the frontend actually runs a BOM, so it is
    exercised here against a stand-in supplier rather than the live APIs."""

    def test_streamed_lookup_emits_progress_and_a_final_summary(self):
        class Stub:
            id = 'digikey'
            name = 'DigiKey'
            configured = True

            def fetch_record(self, part):
                return {
                    'supplier': 'DigiKey',
                    'manufacturerPartNumber': part['mpn'],
                    'leadTime': '10 Weeks',
                    'lifecycle': 'Active',
                    'totalStock': 10000,
                    'currency': 'USD',
                    'variations': [{
                        'supplierPartNumber': 'DK-' + part['mpn'],
                        'stock': 10000,
                        'minimumOrderQuantity': 1,
                        'orderMultiple': 1,
                        'priceBreaks': [{'quantity': 1, 'unitPrice': 0.25}],
                    }],
                }

        original = server.lookup_service.clients
        server.lookup_service.clients = [Stub()]
        try:
            request = urllib.request.Request(
                BASE + '/api/lookup',
                data=json.dumps({
                    'stream': True,
                    'parts': [
                        {'row': 1, 'mpn': 'AAA111', 'quantity': 10},
                        {'row': 2, 'mpn': 'BBB222', 'quantity': 4},
                    ],
                }).encode('utf-8'),
                method='POST',
            )
            request.add_header('Content-Type', 'application/json')
            request.add_header('Accept', 'text/event-stream')
            with urllib.request.urlopen(request, timeout=15) as response:
                self.assertEqual(response.status, 200)
                self.assertIn('text/event-stream', response.headers.get('Content-Type'))
                body = response.read().decode('utf-8')
        finally:
            server.lookup_service.clients = original

        events = []
        for frame in body.split('\n\n'):
            if not frame.strip():
                continue
            lines = frame.split('\n')
            name = next(l[6:].strip() for l in lines if l.startswith('event:'))
            payload = next(l[5:].strip() for l in lines if l.startswith('data:'))
            events.append({'name': name, 'data': json.loads(payload)})

        self.assertEqual(events[0]['name'], 'start')
        self.assertEqual(events[0]['data']['parts'], 2)
        self.assertTrue(any(e['name'] == 'progress' for e in events), 'expected a progress event')

        done = events[-1]
        self.assertEqual(done['name'], 'done')
        self.assertEqual(len(done['data']['rows']), 2)
        self.assertEqual(done['data']['rows'][0]['offers']['digikey']['extendedPrice'], 2.5)
        self.assertEqual(done['data']['summary']['lines'], 2)
        self.assertEqual(done['data']['summary']['supplierTotals']['digikey']['total'], 3.5)
        self.assertEqual(done['data']['summary']['bestMixTotal'], 3.5)


if __name__ == '__main__':
    unittest.main()
