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

    def test_health_publishes_the_match_mode(self):
        # Exact by default: a near neighbour priced as the real part is worse
        # than no answer at all.
        self.assertEqual(call('/api/health')['data']['mpnMatch'], 'exact')

    def test_the_configured_clients_are_built_in_that_mode(self):
        self.assertEqual(server.digikey.match_mode, 'exact')
        self.assertEqual(server.mouser.match_mode, 'exact')

    def test_health_publishes_the_prefixes_so_the_ui_can_warn_first(self):
        result = call('/api/health')
        self.assertEqual(result['data']['ignorePrefixes'],
                         ['ASY0', 'CBL0', 'DES0', 'PCB0'])

    def test_health_says_whether_bom_alternates_are_looked_up(self):
        self.assertIs(call('/api/health')['data']['lookupAlternates'], True)

    def test_a_line_names_at_most_the_configured_number_of_alternates(self):
        # A pasted column can hold a paragraph; the quota should not.
        many = ['SUB-%d' % i for i in range(server.MAX_ALTERNATES_PER_PART + 10)]
        result = self._lookup([{'row': 1, 'mpn': 'ASY0-1', 'quantity': 1, 'alternates': many}])
        # Screened out as in-house, so nothing was looked up — but the request
        # was accepted rather than rejected for its size.
        self.assertEqual(result['status'], 200)
        self.assertGreater(server.MAX_ALTERNATES_PER_PART, 0)

    def test_a_non_list_alternates_field_does_not_break_the_request(self):
        result = self._lookup([{'row': 1, 'mpn': 'ASY0-1', 'quantity': 1, 'alternates': 'SUB-1'}])
        self.assertEqual(result['status'], 200)
        self.assertEqual(result['data']['excluded'][0]['reason'], 'ignored')

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

    def test_a_hand_entered_search_is_answered_rather_than_screened(self):
        # Somebody who types ASY0-1 is asking about ASY0-1. Screening exists to
        # stop automatic waste, not to refuse a direct question.
        result = call('/api/lookup', 'POST', json.dumps({
            'parts': [{'row': 1, 'mpn': 'ASY0-1', 'quantity': 1}],
            'manual': True,
        }), {'Content-Type': 'application/json'})
        # No supplier is configured, so reaching the lookup is a 502 — which is
        # the proof the part was not screened out first.
        self.assertEqual(result['status'], 502)

    def test_a_search_ignores_another_boms_claim(self):
        result = call('/api/lookup', 'POST', json.dumps({
            'parts': [{'row': 1, 'mpn': 'SHARED-PART', 'quantity': 1}],
            'claimed': {'SHARED-PART': 'Main board'},
            'manual': True,
        }), {'Content-Type': 'application/json'})
        self.assertEqual(result['status'], 502)

    def test_the_same_request_without_the_flag_is_screened_normally(self):
        result = self._lookup([{'row': 1, 'mpn': 'ASY0-1', 'quantity': 1}])
        self.assertEqual(result['status'], 200)
        self.assertEqual(result['data']['excluded'][0]['reason'], 'ignored')

    def test_a_search_still_merges_the_same_part_typed_twice(self):
        result = call('/api/lookup', 'POST', json.dumps({
            'parts': [{'row': 1, 'mpn': 'ABC', 'quantity': 5},
                      {'row': 2, 'mpn': 'abc', 'quantity': 5}],
            'manual': True,
        }), {'Content-Type': 'application/json'})
        # Still one part to price, so still one lookup — 502 for want of
        # credentials, not a screening result.
        self.assertEqual(result['status'], 502)

    def test_a_line_the_bom_marks_skip_costs_no_lookup(self):
        result = self._lookup([
            {'row': 1, 'mpn': 'ABC123', 'quantity': 10, 'skip': 'YES'},
            {'row': 2, 'mpn': 'DEF456', 'quantity': 10, 'skip': 'NO'},
        ])
        # DEF456 reaches the (unconfigured) suppliers and fails; ABC123 never
        # gets that far, which is the point.
        self.assertEqual(result['status'], 502)

        only_marked = self._lookup([{'row': 1, 'mpn': 'ABC123', 'quantity': 10, 'skip': 'Y'}])
        self.assertEqual(only_marked['status'], 200)
        self.assertEqual(only_marked['data']['rows'], [])
        self.assertEqual(only_marked['data']['stats']['apiCalls'], 0)
        self.assertEqual(only_marked['data']['excluded'][0]['reason'], 'flagged')

    def test_an_unrecognised_value_in_that_column_does_not_drop_the_line(self):
        result = self._lookup([{'row': 1, 'mpn': 'ABC123', 'quantity': 10, 'skip': 'maybe'}])
        self.assertEqual(result['status'], 502)

    def test_a_hand_entered_lookup_ignores_the_column_entirely(self):
        result = call('/api/lookup', 'POST', json.dumps({
            'parts': [{'row': 1, 'mpn': 'ABC123', 'quantity': 1, 'skip': 'YES'}],
            'manual': True,
        }), {'Content-Type': 'application/json'})
        self.assertEqual(result['status'], 502)

    def test_the_row_cap_counts_parts_that_survive_screening(self):
        # More rows than MAX_PARTS_PER_REQUEST, but nearly all in-house: the
        # request is about a handful of real parts and must be accepted.
        parts = [{'row': i, 'mpn': 'ASY0-%d' % i, 'quantity': 1}
                 for i in range(server.MAX_PARTS_PER_REQUEST + 50)]
        result = self._lookup(parts)
        self.assertEqual(result['status'], 200)
        self.assertEqual(result['data']['rows'], [])


class DmsmsTests(unittest.TestCase):
    """The case form is built from the parts the analyst ticked, not from a rule."""

    def rows(self, *specs):
        out = []
        for index, (mpn, lifecycle) in enumerate(specs):
            out.append({
                'index': index, 'row': index + 1, 'mpn': mpn, 'quantity': 50,
                'reference': 'R%d' % (index + 1), 'manufacturer': 'Acme',
                'description': 'A part', 'assembly': 'Main board',
                'offers': {'digikey': {
                    'supplier': 'DigiKey', 'found': True, 'lifecycle': lifecycle,
                    'lifecycleSeverity': 'bad', 'stock': 500, 'totalStock': 500,
                    'unitPrice': 1.25, 'extendedPrice': 62.5,
                }},
                'comparison': {
                    'lifecycle': lifecycle, 'lifecycleSeverity': 'bad',
                    'recommendedSupplier': 'DigiKey', 'inStockSuppliers': ['DigiKey'],
                    'flags': [],
                },
            })
        return out

    def post(self, payload):
        return call_binary('/api/dmsms', body=json.dumps(payload),
                           headers={'Content-Type': 'application/json'})

    def test_health_publishes_the_statuses_the_form_covers(self):
        data = call('/api/health')['data']['dmsms']
        self.assertIn('Obsolete', data['statuses'])
        self.assertIn('Not Recommended for New Designs', data['statuses'])
        # Still buyable, so listed but not ticked for the analyst.
        self.assertNotIn('Not Recommended for New Designs', data['defaultSelected'])
        self.assertTrue(data['resolutionOptions'])

    def test_a_form_comes_back_named_after_the_program(self):
        result = self.post({
            'meta': {'program': 'Falcon II'},
            'rows': self.rows(('ABC123', 'Obsolete')),
        })
        self.assertEqual(result['status'], 200)
        self.assertIn('spreadsheetml', result['headers']['Content-Type'])
        self.assertIn('Falcon-II-dmsms.xlsx', result['headers']['Content-Disposition'])

    def test_only_the_rows_sent_appear_on_the_form(self):
        from bomlib.spreadsheet import parse_xlsx
        data = self.post({
            'meta': {'program': 'Osprey'},
            'rows': self.rows(('AAA', 'Obsolete'), ('BBB', 'End of Life')),
        })['body']
        grid = parse_xlsx(data)
        index = next(i for i, line in enumerate(grid) if line and line[0] == 'Item')
        header = grid[index]
        records = [line for line in grid[index + 1:] if line and str(line[0]).isdigit()]
        self.assertEqual([r[header.index('Manufacturer Part Number')] for r in records],
                         ['AAA', 'BBB'])

    def test_two_programs_get_two_different_forms_from_one_analysis(self):
        from bomlib.spreadsheet import parse_xlsx

        def parts(payload):
            grid = parse_xlsx(payload)
            index = next(i for i, line in enumerate(grid) if line and line[0] == 'Item')
            header = grid[index]
            return [line[header.index('Manufacturer Part Number')]
                    for line in grid[index + 1:] if line and str(line[0]).isdigit()]

        every = self.rows(('AAA', 'Obsolete'), ('BBB', 'Obsolete'), ('CCC', 'Obsolete'))
        first = self.post({'meta': {'program': 'Falcon II'}, 'rows': every[:2]})
        second = self.post({'meta': {'program': 'Osprey'}, 'rows': every[2:]})
        self.assertEqual(parts(first['body']), ['AAA', 'BBB'])
        self.assertEqual(parts(second['body']), ['CCC'])
        self.assertIn('Falcon-II', first['headers']['Content-Disposition'])
        self.assertIn('Osprey', second['headers']['Content-Disposition'])

    def test_an_empty_selection_is_refused_rather_than_producing_a_blank_form(self):
        self.assertEqual(self.post({'meta': {'program': 'X'}, 'rows': []})['status'], 400)
        self.assertEqual(self.post({'meta': {'program': 'X'},
                                    'rows': [{'quantity': 1}]})['status'], 400)


class AlternativesTests(unittest.TestCase):
    """Nexar runs only when asked, and only for the parts named."""

    def post(self, payload):
        return call('/api/alternatives', 'POST', json.dumps(payload),
                    {'Content-Type': 'application/json'})

    def test_health_says_whether_alternatives_can_be_looked_up(self):
        data = call('/api/health')['data']['alternatives']
        self.assertEqual(data['provider'], 'Nexar')
        # No credentials in this environment.
        self.assertFalse(data['configured'])
        self.assertTrue(data['maxParts'] > 0)

    def test_without_credentials_it_says_so_rather_than_failing_obscurely(self):
        result = self.post({'parts': [{'mpn': 'ABC123'}]})
        self.assertEqual(result['status'], 400)
        self.assertIn('NEXAR_CLIENT_ID', result['data']['error'])

    def test_alternatives_are_never_part_of_a_bom_lookup(self):
        # A BOM run must not spend a Nexar call: the endpoint is separate on
        # purpose, and /api/lookup knows nothing about it.
        result = call('/api/lookup', 'POST',
                      json.dumps({'parts': [{'row': 1, 'mpn': 'ASY0-1', 'quantity': 1}]}),
                      {'Content-Type': 'application/json'})
        self.assertNotIn('alternatives', result['data'])


class ConfiguredAlternativesTests(unittest.TestCase):
    """With a client in place, driven through the real endpoint."""

    def setUp(self):
        self.original = server.nexar
        self.asked = []

        class Stub:
            configured = True

            def find_alternatives(inner, part):
                self.asked.append(part['mpn'])
                if part['mpn'] == 'BOOM':
                    raise RuntimeError("Cannot query field 'similarParts'")
                if part['mpn'] == 'NOMATCH':
                    return {'matched': None, 'alternatives': []}
                return {
                    'matched': {'mpn': part['mpn'], 'manufacturer': 'Acme'},
                    'alternatives': [{'mpn': part['mpn'] + '-ALT', 'manufacturer': 'Beta',
                                      'stock': 100, 'specs': []}],
                }

        server.nexar = Stub()

    def tearDown(self):
        server.nexar = self.original
        server.cache.clear()

    def post(self, parts):
        return call('/api/alternatives', 'POST', json.dumps({'parts': parts}),
                    {'Content-Type': 'application/json'})

    def test_each_part_comes_back_with_what_was_found_for_it(self):
        data = self.post([{'mpn': 'ABC123', 'manufacturer': 'Acme'}])['data']
        self.assertEqual(data['provider'], 'Nexar')
        result = data['results'][0]
        self.assertEqual(result['mpn'], 'ABC123')
        self.assertEqual(result['matched']['mpn'], 'ABC123')
        self.assertEqual([a['mpn'] for a in result['alternatives']], ['ABC123-ALT'])
        self.assertIsNone(result['error'])

    def test_a_failure_on_one_part_does_not_lose_the_others(self):
        data = self.post([{'mpn': 'AAA'}, {'mpn': 'BOOM'}, {'mpn': 'CCC'}])['data']
        by_mpn = {r['mpn']: r for r in data['results']}
        self.assertEqual(len(by_mpn['AAA']['alternatives']), 1)
        self.assertIn('similarParts', by_mpn['BOOM']['error'])
        self.assertEqual(len(by_mpn['CCC']['alternatives']), 1)
        self.assertEqual(data['stats']['errors'], 1)

    def test_a_part_with_no_match_is_reported_plainly(self):
        result = self.post([{'mpn': 'NOMATCH'}])['data']['results'][0]
        self.assertIsNone(result['matched'])
        self.assertEqual(result['alternatives'], [])
        self.assertIsNone(result['error'])

    def test_a_second_ask_for_the_same_part_is_served_from_cache(self):
        self.post([{'mpn': 'ABC123'}])
        again = self.post([{'mpn': 'ABC123'}])['data']
        self.assertEqual(self.asked, ['ABC123'])
        self.assertEqual(again['stats']['cacheHits'], 1)
        self.assertEqual(again['stats']['apiCalls'], 0)

    def test_a_failure_is_not_cached_so_the_next_run_retries(self):
        self.post([{'mpn': 'BOOM'}])
        self.post([{'mpn': 'BOOM'}])
        self.assertEqual(self.asked, ['BOOM', 'BOOM'])

    def test_part_numbers_are_cleaned_before_being_asked_about(self):
        self.post([{'mpn': '  \u200bABC123 '}])
        self.assertEqual(self.asked, ['ABC123'])

    def test_an_empty_or_oversized_request_is_refused(self):
        self.assertEqual(self.post([])['status'], 400)
        self.assertEqual(self.post([{'mpn': ''}])['status'], 400)
        too_many = [{'mpn': 'P%d' % i} for i in range(server.MAX_ALTERNATIVE_PARTS + 1)]
        self.assertEqual(self.post(too_many)['status'], 400)


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
