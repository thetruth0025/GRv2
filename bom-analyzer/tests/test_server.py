import json
import os
import threading
import unittest
import urllib.error
import urllib.request

# Keep the test run away from the developer's real credentials and cache file.
os.environ['CACHE_FILE'] = 'none'
for name in ('DIGIKEY_CLIENT_ID', 'DIGIKEY_CLIENT_SECRET', 'MOUSER_API_KEY',
             'TRUSTEDPARTS_API_KEY'):
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


def call_binary(path, method='POST', body=None, headers=None):
    """For endpoints that answer with a file rather than JSON."""
    data = body.encode('utf-8') if isinstance(body, str) else body
    request = urllib.request.Request(BASE + path, data=data, method=method)
    for key, value in (headers or {}).items():
        request.add_header(key, value)
    try:
        with urllib.request.urlopen(request, timeout=10) as response:
            return {'status': response.status, 'body': response.read(),
                    'headers': dict(response.headers)}
    except urllib.error.HTTPError as err:
        return {'status': err.code, 'body': err.read(), 'headers': dict(err.headers)}


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
        self.assertEqual([s['id'] for s in result['data']['suppliers']],
                         ['digikey', 'mouser', 'trustedparts'])
        # No credentials are set in this environment.
        self.assertTrue(all(s['configured'] is False for s in result['data']['suppliers']))
        # The frontend needs to know which supplier aggregates other distributors.
        aggregators = [s['id'] for s in result['data']['suppliers'] if s.get('aggregator')]
        self.assertEqual(aggregators, ['trustedparts'])


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


class ScreeningTests(unittest.TestCase):
    """In-house and repeated part numbers must not reach a supplier."""

    def test_health_publishes_the_prefixes_so_the_ui_can_warn_first(self):
        result = call('/api/health')
        self.assertEqual(result['data']['ignorePrefixes'],
                         ['ASY0', 'CBL0', 'DES0', 'PCB0'])

    def _lookup(self, parts, claimed=None):
        payload = {'parts': parts}
        if claimed is not None:
            payload['claimed'] = claimed
        return call('/api/lookup', 'POST', json.dumps(payload),
                    {'Content-Type': 'application/json'})

    def test_a_bom_of_only_in_house_numbers_costs_no_lookup_at_all(self):
        result = self._lookup([
            {'row': 1, 'mpn': 'ASY0-1', 'quantity': 1},
            {'row': 2, 'mpn': 'CBL0-2', 'quantity': 1},
            {'row': 3, 'mpn': 'DES0-3', 'quantity': 1},
            {'row': 4, 'mpn': 'PCB0-4', 'quantity': 1},
        ])
        # No supplier is configured here, so a lookup that was attempted would
        # have failed outright — a 200 with no rows proves none was attempted.
        self.assertEqual(result['status'], 200)
        self.assertEqual(result['data']['rows'], [])
        self.assertEqual(result['data']['stats']['apiCalls'], 0)
        self.assertEqual([e['reason'] for e in result['data']['excluded']],
                         ['ignored'] * 4)

    def test_a_part_claimed_by_another_bom_is_reported_not_looked_up(self):
        result = self._lookup(
            [{'row': 1, 'mpn': 'PCB0-9', 'quantity': 1},
             {'row': 2, 'mpn': 'SHARED-PART', 'quantity': 4}],
            claimed={'SHARED-PART': 'Main board'},
        )
        self.assertEqual(result['status'], 200)
        reasons = {e['mpn']: e for e in result['data']['excluded']}
        self.assertEqual(reasons['SHARED-PART']['reason'], 'duplicate')
        self.assertIn('Main board', reasons['SHARED-PART']['detail'])

    def test_a_malformed_claim_map_is_ignored_rather_than_failing_the_request(self):
        result = call('/api/lookup', 'POST',
                      json.dumps({'parts': [{'row': 1, 'mpn': 'ASY0-1'}], 'claimed': 'nonsense'}),
                      {'Content-Type': 'application/json'})
        self.assertEqual(result['status'], 200)

    def test_the_row_cap_counts_parts_that_survive_screening(self):
        # More rows than MAX_PARTS_PER_REQUEST, but nearly all in-house: the
        # request is about a handful of real parts and must be accepted.
        parts = [{'row': i, 'mpn': 'ASY0-%d' % i, 'quantity': 1}
                 for i in range(server.MAX_PARTS_PER_REQUEST + 50)]
        result = self._lookup(parts)
        self.assertEqual(result['status'], 200)
        self.assertEqual(result['data']['rows'], [])


class ReportTests(unittest.TestCase):
    """The workbook is built from results the client already holds."""

    def book(self, name='Main board'):
        offer = {
            'supplier': 'DigiKey', 'found': True, 'unitPrice': 0.05, 'extendedPrice': 5.0,
            'stock': 5000, 'stockSufficient': True, 'lifecycle': 'Active',
            'lifecycleSeverity': 'ok', 'leadTimeText': '10 weeks', 'leadTimeDays': 70,
            'currency': 'USD', 'orderQuantity': 100, 'supplierPartNumber': 'DK-1',
        }
        return {
            'name': name,
            'suppliers': [{'id': 'digikey', 'name': 'DigiKey'}],
            'rows': [{
                'index': 0, 'row': 1, 'mpn': 'RC0603FR-0710KL', 'quantity': 100,
                'reference': 'R1', 'manufacturer': 'Yageo', 'description': '10k 1%',
                'offers': {'digikey': offer},
                'comparison': {
                    'bestPriceSupplier': 'DigiKey', 'bestPrice': 5.0,
                    'bestLeadTimeSuppliers': ['DigiKey'], 'bestLeadTimeDays': 0,
                    'inStockSuppliers': ['DigiKey'], 'lifecycle': 'Active',
                    'lifecycleSeverity': 'ok', 'recommendedSupplier': 'DigiKey', 'flags': [],
                },
            }],
            'summary': {
                'lines': 1, 'totalQuantity': 100, 'currency': 'USD', 'bestMixTotal': 5.0,
                'bestMixLines': 1, 'notFoundLines': 0,
                'supplierTotals': {'digikey': {
                    'id': 'digikey', 'name': 'DigiKey', 'total': 5.0, 'linesPriced': 1,
                    'linesMissing': 0, 'linesShort': 0, 'complete': True,
                }},
                'cheapestSingleSource': 'digikey', 'mixSavings': 0.0,
            },
            'excluded': [{'row': 2, 'mpn': 'ASY0-1', 'quantity': 1, 'reason': 'ignored',
                          'detail': 'In-house part number (starts with ASY0)'}],
        }

    def post(self, books):
        return call_binary('/api/report', body=json.dumps({'books': books}),
                           headers={'Content-Type': 'application/json'})

    def test_a_workbook_comes_back_as_a_named_download(self):
        result = self.post([self.book()])
        self.assertEqual(result['status'], 200)
        self.assertIn('spreadsheetml', result['headers']['Content-Type'])
        self.assertIn('Main-board-report.xlsx', result['headers']['Content-Disposition'])

    def test_the_sheets_lead_with_the_report_and_include_the_skipped_lines(self):
        from bomlib.spreadsheet import parse_xlsx
        data = self.post([self.book()])['body']
        self.assertEqual(parse_xlsx(data)[0][0], 'BOM Supplier Report')
        self.assertEqual(parse_xlsx(data, 'Parts')[1][1], 'RC0603FR-0710KL')
        self.assertEqual(parse_xlsx(data, 'Skipped')[1][1], 'ASY0-1')

    def test_several_boms_share_one_workbook_with_prefixed_tabs(self):
        from bomlib.spreadsheet import parse_xlsx
        result = self.post([self.book('Alpha'), self.book('Beta')])
        self.assertIn('all-boms-report.xlsx', result['headers']['Content-Disposition'])
        data = result['body']
        self.assertEqual(parse_xlsx(data, 'Alpha Report')[0][0], 'BOM Supplier Report')
        self.assertEqual(parse_xlsx(data, 'Beta Parts')[1][1], 'RC0603FR-0710KL')

    def test_an_empty_request_is_refused_rather_than_producing_a_blank_file(self):
        self.assertEqual(self.post([])['status'], 400)
        self.assertEqual(self.post([{'name': 'x', 'rows': []}])['status'], 400)

    def test_a_part_number_that_looks_like_a_formula_stays_text(self):
        from bomlib.spreadsheet import parse_xlsx
        book = self.book()
        book['rows'][0]['mpn'] = '=cmd|calc'
        data = self.post([book])['body']
        # Written as an inline string, so a spreadsheet shows it rather than
        # handing it to the formula engine.
        self.assertEqual(parse_xlsx(data, 'Parts')[1][1], '=cmd|calc')

    def test_a_filename_cannot_be_smuggled_through_the_bom_name(self):
        result = self.post([self.book('../../etc/passwd"; drop')])
        disposition = result['headers']['Content-Disposition']
        self.assertNotIn('"', disposition[len('attachment; filename='):].strip('"'))
        self.assertNotIn('/', disposition)
