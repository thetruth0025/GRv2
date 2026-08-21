import unittest

from bomlib.http_client import HttpError
from bomlib.lookup import LookupService, summarize_bom
from bomlib.normalize import Lifecycle, record_to_offer
from bomlib.report import build_distributor_rows, has_distributor_detail
from bomlib.trustedparts import TrustedPartsClient, stock_of

# Shaped exactly like the v2 OpenAPI schema: ApiResponse → PartResults →
# Distributors → DistributorResults.
RESPONSE = {
    'Messages': [],
    'ErrorMessage': None,
    'PartResults': [{
        'PartNumber': 'STM32F103C8T6',
        'Manufacturer': 'STMicroelectronics',
        'ManufacturerId': 42,
        'ProductUrl': 'https://www.trustedparts.com/en/part/stm/STM32F103C8T6',
        'IsAffectedByTariff': True,
        'LifecycleRisk': None,
        'SupplyChainRisk': None,
        'Distributors': [
            {
                'Id': 1,
                'Name': 'Arrow Electronics',
                'DistributorResults': [{
                    'Description': 'ARM Cortex-M3 MCU 64KB LQFP48',
                    'DistributorPartNumber': 'ARW-STM32F103C8T6',
                    'Compliance': {'RoHS': [{'Region': 'EU', 'IsCompliant': True,
                                             'Description': 'RoHS Compliant'}]},
                    'Stock': {'QuantityOnHand': 12400, 'Availability': 'In Stock'},
                    'Links': [
                        {'Type': 'datasheet', 'Url': 'https://example.com/stm32.pdf'},
                        {'Type': 'distributor', 'Url': 'https://arrow.com/p/stm32'},
                    ],
                    'Pricing': {
                        'CurrencyCode': 'USD',
                        'MinimumQuantity': 1,
                        'QuantityMultiple': 1,
                        'Prices': [
                            {'Quantity': 1, 'Amount': 3.10, 'FormattedAmount': '$3.10'},
                            {'Quantity': 100, 'Amount': 2.44, 'FormattedAmount': '$2.44'},
                        ],
                    },
                    'Packaging': [{'PackageType': 'Tray', 'MinimumOrderQuantity': 1}],
                }],
            },
            {
                'Id': 2,
                'Name': 'Future Electronics',
                'DistributorResults': [{
                    'Description': 'ARM Cortex-M3 MCU 64KB LQFP48',
                    'DistributorPartNumber': 'FUT-STM32F103C8T6',
                    'Stock': {'QuantityOnHand': 300, 'Availability': 'In Stock'},
                    'Pricing': {
                        'CurrencyCode': 'USD',
                        'MinimumQuantity': 1,
                        'QuantityMultiple': 1,
                        'Prices': [{'Quantity': 1, 'Amount': 2.20}],
                    },
                    'Packaging': [],
                }],
            },
            {
                'Id': 3,
                'Name': 'TTI Inc',
                'DistributorResults': [{
                    'DistributorPartNumber': 'TTI-STM32F103C8T6',
                    'Stock': {'QuantityOnHand': None, 'Availability': 'In Stock'},
                    'Pricing': {
                        'CurrencyCode': 'USD',
                        'MinimumQuantity': 500,
                        'QuantityMultiple': 500,
                        'Prices': [{'Quantity': 500, 'Amount': 1.90}],
                    },
                }],
            },
        ],
    }],
}

PART = {'mpn': 'STM32F103C8T6', 'manufacturer': 'STMicroelectronics', 'quantity': 1000}


class RecordTests(unittest.TestCase):
    def setUp(self):
        self.client = TrustedPartsClient(api_key='key')

    def test_every_distributor_becomes_a_variation(self):
        records = self.client.to_records(RESPONSE, [PART])
        record = records['STM32F103C8T6']
        self.assertTrue(record['aggregator'])
        self.assertEqual(record['manufacturerPartNumber'], 'STM32F103C8T6')
        names = sorted(v['distributor'] for v in record['variations'])
        self.assertEqual(names, ['Arrow Electronics', 'Future Electronics', 'TTI Inc'])

    def test_no_lead_time_is_claimed_because_the_api_reports_none(self):
        record = self.client.to_records(RESPONSE, [PART])['STM32F103C8T6']
        offer = record_to_offer(record, PART)
        self.assertIsNone(record['leadTime'])
        self.assertIsNone(offer['leadTimeDays'])
        self.assertIsNone(offer['leadTimeText'])

    def test_the_tariff_flag_is_carried_through_to_the_offer(self):
        offer = record_to_offer(self.client.to_records(RESPONSE, [PART])['STM32F103C8T6'], PART)
        self.assertTrue(offer['affectedByTariff'])

    def test_the_column_quotes_the_cheapest_distributor_that_can_cover_the_line(self):
        offer = record_to_offer(self.client.to_records(RESPONSE, [PART])['STM32F103C8T6'], PART)
        # 1000 pieces: Future stocks only 300, TTI's depth is undisclosed, so
        # Arrow is the one that can actually ship it.
        self.assertEqual(offer['distributor'], 'Arrow Electronics')
        self.assertEqual(offer['unitPrice'], 2.44)
        self.assertEqual(offer['extendedPrice'], 2440.0)
        self.assertTrue(offer['stockSufficient'])

    def test_every_distributor_is_priced_and_kept_for_the_detail_view(self):
        offer = record_to_offer(self.client.to_records(RESPONSE, [PART])['STM32F103C8T6'], PART)
        self.assertEqual(offer['distributorCount'], 3)
        names = [d['distributor'] for d in offer['distributorOffers']]
        self.assertEqual(sorted(names), ['Arrow Electronics', 'Future Electronics', 'TTI Inc'])
        # Ordered with the one that covers the line first.
        self.assertEqual(names[0], 'Arrow Electronics')
        future = next(d for d in offer['distributorOffers'] if d['distributor'] == 'Future Electronics')
        self.assertFalse(future['stockSufficient'])
        self.assertEqual(future['unitPrice'], 2.20)

    def test_a_smaller_quantity_can_change_which_distributor_wins(self):
        record = self.client.to_records(RESPONSE, [PART])['STM32F103C8T6']
        offer = record_to_offer(record, {'mpn': PART['mpn'], 'quantity': 100})
        # At 100 pieces Future covers it and is cheapest.
        self.assertEqual(offer['distributor'], 'Future Electronics')
        self.assertEqual(offer['extendedPrice'], 220.0)

    def test_a_minimum_order_is_respected_when_pricing_a_distributor(self):
        record = self.client.to_records(RESPONSE, [PART])['STM32F103C8T6']
        offer = record_to_offer(record, PART)
        tti = next(d for d in offer['distributorOffers'] if d['distributor'] == 'TTI Inc')
        self.assertEqual(tti['minimumOrderQuantity'], 500)
        self.assertEqual(tti['orderQuantity'], 1000)

    def test_rohs_and_links_are_read_from_the_nested_shapes(self):
        record = self.client.to_records(RESPONSE, [PART])['STM32F103C8T6']
        self.assertEqual(record['rohs'], 'RoHS Compliant')
        self.assertEqual(record['datasheetUrl'], 'https://example.com/stm32.pdf')
        self.assertEqual(record['productUrl'],
                         'https://www.trustedparts.com/en/part/stm/STM32F103C8T6')


class StockTests(unittest.TestCase):
    def test_a_disclosed_quantity_is_used_directly(self):
        self.assertEqual(stock_of({'Stock': {'QuantityOnHand': 4200}}), 4200)

    def test_an_undisclosed_quantity_stays_unknown_rather_than_zero(self):
        # "In Stock" with no number means stocked at an undisclosed depth.
        self.assertIsNone(stock_of({'Stock': {'QuantityOnHand': None, 'Availability': 'In Stock'}}))

    def test_an_explicit_none_availability_reads_as_zero(self):
        self.assertEqual(stock_of({'Stock': {'QuantityOnHand': None, 'Availability': 'None'}}), 0)
        self.assertEqual(
            stock_of({'Stock': {'QuantityOnHand': None, 'Availability': 'Out of Stock'}}), 0)

    def test_a_number_inside_availability_prose_is_recovered(self):
        self.assertEqual(
            stock_of({'Stock': {'QuantityOnHand': None, 'Availability': '1,250 In Stock'}}), 1250)


class LifecycleRiskTests(unittest.TestCase):
    """A risk grade is not a lifecycle status and must never be rendered as one."""

    def setUp(self):
        self.client = TrustedPartsClient(api_key='key')

    def _record_with_risk(self, risk):
        import copy
        payload = copy.deepcopy(RESPONSE)
        payload['PartResults'][0]['LifecycleRisk'] = risk
        return self.client.to_records(payload, [PART])['STM32F103C8T6']

    def test_a_risk_grade_does_not_become_a_lifecycle_status(self):
        record = self._record_with_risk('Low')
        self.assertIsNone(record['lifecycle'])
        self.assertEqual(record['lifecycleRisk'], 'Low')
        offer = record_to_offer(record, PART)
        self.assertEqual(offer['lifecycle'], Lifecycle.UNKNOWN)
        self.assertEqual(offer['lifecycleRisk'], 'Low')

    def test_an_actual_status_is_promoted(self):
        record = self._record_with_risk('Obsolete')
        offer = record_to_offer(record, PART)
        self.assertEqual(offer['lifecycle'], Lifecycle.OBSOLETE)

    def test_a_withheld_risk_leaves_lifecycle_unknown(self):
        offer = record_to_offer(self._record_with_risk(None), PART)
        self.assertEqual(offer['lifecycle'], Lifecycle.UNKNOWN)


class RequestTests(unittest.TestCase):
    def test_the_request_body_matches_the_documented_schema(self):
        client = TrustedPartsClient(api_key='secret', currency='EUR', country='DE')
        captured = {}

        def fake_request(url, method=None, headers=None, body=None, **kwargs):
            import json
            captured['url'] = url
            captured['headers'] = headers
            captured['body'] = json.loads(body)
            return {'data': RESPONSE}

        import bomlib.trustedparts as module
        original = module.request_json
        module.request_json = fake_request
        try:
            client.search([PART, {'mpn': 'LM358DR', 'quantity': 10}])
        finally:
            module.request_json = original

        self.assertEqual(captured['url'], 'https://api.trustedparts.com/v2/search')
        self.assertEqual(captured['headers']['X-Api-Key'], 'secret')
        body = captured['body']
        self.assertEqual([q['SearchToken'] for q in body['Queries']],
                         ['STM32F103C8T6', 'LM358DR'])
        self.assertEqual(body['Queries'][0]['Manufacturers'], ['STMicroelectronics'])
        self.assertEqual(body['CurrencyCode'], 'EUR')
        self.assertEqual(body['CountryCode'], 'DE')
        # A BOM lookup wants the part asked for, not near misses.
        self.assertTrue(body['ExactMatch'])

    def test_a_body_level_error_raises_rather_than_reading_as_empty(self):
        client = TrustedPartsClient(api_key='bad')
        import bomlib.trustedparts as module
        original = module.request_json
        module.request_json = lambda *a, **k: {'data': {'ErrorMessage': 'Invalid API key'}}
        try:
            with self.assertRaisesRegex(HttpError, 'Invalid API key'):
                client.search([PART])
        finally:
            module.request_json = original

    def test_tokens_shorter_than_the_api_minimum_are_dropped(self):
        client = TrustedPartsClient(api_key='key')
        self.assertEqual(client.search([{'mpn': 'X', 'quantity': 1}]), {})


class BatchTests(unittest.TestCase):
    """The aggregator is queried in batches, unlike the per-part suppliers."""

    class BatchStub:
        id = 'trustedparts'
        name = 'TrustedParts'
        configured = True
        batch_size = 50

        def __init__(self):
            self.calls = []

        def fetch_records(self, parts):
            self.calls.append([p['mpn'] for p in parts])
            client = TrustedPartsClient(api_key='key')
            return {p['mpn']: client.to_records(RESPONSE, [p]).get(p['mpn']) for p in parts}

    def test_many_parts_cost_one_request(self):
        stub = self.BatchStub()
        service = LookupService(clients=[stub], cache=None)
        parts = [{'row': i, 'mpn': 'STM32F103C8T6', 'quantity': 100} for i in range(1, 6)]
        result = service.lookup_parts(parts)
        # Five identical lines dedupe to one part, in one batched request.
        self.assertEqual(len(stub.calls), 1)
        self.assertEqual(result['stats']['apiCalls'], 1)
        self.assertEqual(len(result['rows']), 5)

    def test_a_batch_larger_than_the_limit_is_chunked(self):
        stub = self.BatchStub()
        stub.batch_size = 2
        service = LookupService(clients=[stub], cache=None)
        parts = [{'row': i, 'mpn': 'MPN%03d' % i, 'quantity': 1} for i in range(5)]
        service.lookup_parts(parts)
        self.assertEqual([len(c) for c in stub.calls], [2, 2, 1])

    def test_a_failed_batch_marks_every_part_in_it_without_caching(self):
        class Boom:
            id = 'trustedparts'
            name = 'TrustedParts'
            configured = True
            batch_size = 50
            attempts = 0

            def fetch_records(self, parts):
                Boom.attempts += 1
                raise RuntimeError('HTTP 429 from api.trustedparts.com')

        from bomlib.cache import PartCache
        cache = PartCache(ttl_seconds=60, path=None)
        service = LookupService(clients=[Boom()], cache=cache)
        parts = [{'row': 1, 'mpn': 'AAA111', 'quantity': 1}, {'row': 2, 'mpn': 'BBB222', 'quantity': 1}]
        result = service.lookup_parts(parts)
        for row in result['rows']:
            self.assertTrue(row['offers']['trustedparts']['error'])
            self.assertIn('429', row['offers']['trustedparts']['reason'])
        # Not cached, so a later run retries.
        service.lookup_parts(parts)
        self.assertEqual(Boom.attempts, 2)

    def test_progress_is_reported_once_per_part_even_when_batched(self):
        stub = self.BatchStub()
        service = LookupService(clients=[stub], cache=None)
        parts = [{'row': i, 'mpn': 'MPN%03d' % i, 'quantity': 1} for i in range(4)]
        seen = []
        service.lookup_parts(parts, on_progress=lambda p: seen.append(p['completed']))
        self.assertEqual(sorted(seen), [1, 2, 3, 4])


class ExportTests(unittest.TestCase):
    def _result(self):
        stub = BatchTests.BatchStub()
        service = LookupService(clients=[stub], cache=None)
        result = service.lookup_parts([{'row': 2, 'mpn': 'STM32F103C8T6', 'quantity': 1000}])
        return result, summarize_bom(result['rows'], result['suppliers'])

    def test_a_distributor_breakdown_is_detected(self):
        result, _ = self._result()
        self.assertTrue(has_distributor_detail(result))

    def test_the_breakdown_has_one_row_per_distributor_offer(self):
        result, summary = self._result()
        rows = build_distributor_rows(result, summary)
        self.assertEqual(rows[0][:5], ['Row', 'Part Number', 'Quantity', 'Via', 'Distributor'])
        self.assertEqual(len(rows), 4)  # header + three distributors
        names = sorted(r[4] for r in rows[1:])
        self.assertEqual(names, ['Arrow Electronics', 'Future Electronics', 'TTI Inc'])
        for row in rows[1:]:
            self.assertEqual(row[1], 'STM32F103C8T6')
            self.assertEqual(row[3], 'TrustedParts')

    def test_a_bom_without_an_aggregator_gets_no_breakdown(self):
        class Single:
            id = 'digikey'
            name = 'DigiKey'
            configured = True

            def fetch_record(self, part):
                return {
                    'supplier': 'DigiKey',
                    'manufacturerPartNumber': part['mpn'],
                    'totalStock': 10,
                    'currency': 'USD',
                    'variations': [{'stock': 10, 'minimumOrderQuantity': 1, 'orderMultiple': 1,
                                    'priceBreaks': [{'quantity': 1, 'unitPrice': 1.0}]}],
                }

        service = LookupService(clients=[Single()], cache=None)
        result = service.lookup_parts([{'row': 1, 'mpn': 'ABC', 'quantity': 1}])
        self.assertFalse(has_distributor_detail(result))


if __name__ == '__main__':
    unittest.main()
