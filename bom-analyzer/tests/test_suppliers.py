import unittest

from bomlib.cache import PartCache
from bomlib.digikey import DigiKeyClient
from bomlib.http_client import HttpError
from bomlib.lookup import LookupService, summarize_bom
from bomlib.mouser import MouserClient
from bomlib.normalize import Lifecycle, record_to_offer

# Fixtures follow the shapes DigiKey Product Information V4 and the Mouser
# Search API v1 actually return, trimmed to the fields this app reads.

DIGIKEY_RESPONSE = {
    'ProductsCount': 1,
    'ExactMatches': [],
    'Products': [{
        'Description': {
            'ProductDescription': 'RES SMD 10K OHM 1% 1/10W 0603',
            'DetailedDescription': '10 kOhms ±1% 0.1W, 1/10W Chip Resistor 0603',
        },
        'Manufacturer': {'Id': 311, 'Name': 'YAGEO'},
        'ManufacturerProductNumber': 'RC0603FR-0710KL',
        'ProductUrl': 'https://www.digikey.com/en/products/detail/yageo/RC0603FR-0710KL/727385',
        'DatasheetUrl': 'https://www.yageo.com/upload/media/product/RC_L_0.pdf',
        'QuantityAvailable': 1250000,
        'ManufacturerLeadWeeks': '10 Weeks',
        'ProductStatus': {'Id': 0, 'Status': 'Active'},
        'Classifications': {'RohsStatus': 'ROHS3 Compliant'},
        'NormallyStocking': True,
        'ProductVariations': [
            {
                'DigiKeyProductNumber': '311-10.0KHRCT-ND',
                'PackageType': {'Id': 2, 'Name': 'Cut Tape (CT)'},
                'QuantityAvailableforPackageType': 5000,
                'MinimumOrderQuantity': 1,
                'StandardPackage': 1,
                'MarketPlace': False,
                'StandardPricing': [
                    {'BreakQuantity': 1, 'UnitPrice': 0.1, 'TotalPrice': 0.1},
                    {'BreakQuantity': 10, 'UnitPrice': 0.038, 'TotalPrice': 0.38},
                    {'BreakQuantity': 100, 'UnitPrice': 0.0156, 'TotalPrice': 1.56},
                    {'BreakQuantity': 1000, 'UnitPrice': 0.00784, 'TotalPrice': 7.84},
                ],
            },
            {
                'DigiKeyProductNumber': '311-10.0KHRTR-ND',
                'PackageType': {'Id': 1, 'Name': 'Tape & Reel (TR)'},
                'QuantityAvailableforPackageType': 1245000,
                'MinimumOrderQuantity': 5000,
                'StandardPackage': 5000,
                'MarketPlace': False,
                'StandardPricing': [{'BreakQuantity': 5000, 'UnitPrice': 0.00518, 'TotalPrice': 25.9}],
            },
        ],
    }],
}

MOUSER_RESPONSE = {
    'Errors': [],
    'SearchResults': {
        'NumberOfResult': 1,
        'Parts': [{
            'Availability': '128,430 In Stock',
            'AvailabilityInStock': '128430',
            'FactoryStock': '500000',
            'DataSheetUrl': 'https://www.mouser.com/datasheet/2/447/RC_L-1666295.pdf',
            'Description': 'Thick Film Resistors - SMD 10 KOhms 1% 0603',
            'LeadTime': '84 Days',
            'LifecycleStatus': '',
            'ProductStatus': 'New at Mouser',
            'Manufacturer': 'YAGEO',
            'ManufacturerPartNumber': 'RC0603FR-0710KL',
            'Min': '1',
            'Mult': '1',
            'MouserPartNumber': '603-RC0603FR-0710KL',
            'ProductDetailUrl': 'https://www.mouser.com/ProductDetail/603-RC0603FR-0710KL',
            'ROHSStatus': 'RoHS Compliant',
            'Reeling': False,
            'PriceBreaks': [
                {'Quantity': 1, 'Price': '$0.10', 'Currency': 'USD'},
                {'Quantity': 100, 'Price': '$0.017', 'Currency': 'USD'},
                {'Quantity': 1000, 'Price': '$0.009', 'Currency': 'USD'},
            ],
        }],
    },
}

PART = {'mpn': 'RC0603FR-0710KL', 'manufacturer': 'Yageo', 'quantity': 500}


class DigiKeyTests(unittest.TestCase):
    def setUp(self):
        self.client = DigiKeyClient(client_id='id', client_secret='secret')

    def test_v4_response_becomes_a_priced_offer(self):
        record = self.client.to_record(DIGIKEY_RESPONSE, PART)
        self.assertIsNotNone(record)
        offer = record_to_offer(record, PART)

        self.assertEqual(offer['supplier'], 'DigiKey')
        self.assertEqual(offer['manufacturerPartNumber'], 'RC0603FR-0710KL')
        self.assertEqual(offer['manufacturer'], 'YAGEO')
        self.assertEqual(offer['lifecycle'], Lifecycle.ACTIVE)
        self.assertEqual(offer['leadTimeDays'], 70)
        self.assertEqual(offer['leadTimeText'], '10 weeks')
        self.assertEqual(offer['rohs'], 'ROHS3 Compliant')
        self.assertEqual(offer['description'], 'RES SMD 10K OHM 1% 1/10W 0603')
        self.assertTrue(offer['exactMatch'])

        # 500 pieces fit cut tape, which is stocked and needs no 5000 minimum.
        self.assertEqual(offer['supplierPartNumber'], '311-10.0KHRCT-ND')
        self.assertEqual(offer['orderQuantity'], 500)
        self.assertEqual(offer['unitPrice'], 0.0156)
        self.assertEqual(offer['extendedPrice'], 7.8)
        self.assertTrue(offer['stockSufficient'])

    def test_order_beyond_cut_tape_stock_moves_to_the_reel(self):
        record = self.client.to_record(DIGIKEY_RESPONSE, PART)
        offer = record_to_offer(record, {'mpn': PART['mpn'], 'quantity': 20000})
        self.assertEqual(offer['supplierPartNumber'], '311-10.0KHRTR-ND')
        self.assertEqual(offer['minimumOrderQuantity'], 5000)
        self.assertEqual(offer['orderQuantity'], 20000)
        self.assertEqual(offer['unitPrice'], 0.00518)

    def test_v3_shaped_product_still_parses(self):
        legacy = {'Products': [{
            'ManufacturerPartNumber': 'RC0603FR-0710KL',
            'Manufacturer': {'Value': 'YAGEO'},
            'ProductDescription': 'RES SMD 10K OHM',
            'DigiKeyPartNumber': '311-10.0KHRCT-ND',
            'QuantityAvailable': 5000,
            'MinimumOrderQuantity': 1,
            'ProductStatus': 'Active',
            'StandardPricing': [{'BreakQuantity': 1, 'UnitPrice': 0.1}],
        }]}
        offer = record_to_offer(self.client.to_record(legacy, PART), {'mpn': PART['mpn'], 'quantity': 10})
        self.assertEqual(offer['supplierPartNumber'], '311-10.0KHRCT-ND')
        self.assertEqual(offer['manufacturer'], 'YAGEO')
        self.assertEqual(offer['unitPrice'], 0.1)


class MouserTests(unittest.TestCase):
    def setUp(self):
        self.client = MouserClient(api_key='key')

    def test_response_becomes_a_priced_offer(self):
        record = self.client.to_record(MOUSER_RESPONSE, PART)
        self.assertIsNotNone(record)
        offer = record_to_offer(record, PART)

        self.assertEqual(offer['supplier'], 'Mouser')
        self.assertEqual(offer['supplierPartNumber'], '603-RC0603FR-0710KL')
        self.assertEqual(offer['stock'], 128430)
        self.assertEqual(offer['factoryStock'], 500000)
        self.assertEqual(offer['leadTimeDays'], 84)
        self.assertEqual(offer['unitPrice'], 0.017)
        self.assertEqual(offer['extendedPrice'], 8.5)
        self.assertTrue(offer['stockSufficient'])
        # LifecycleStatus was blank, so the Mouser catalog status is used.
        self.assertEqual(offer['lifecycle'], Lifecycle.NEW)

    def test_unrelated_result_is_rejected_rather_than_reported_as_a_match(self):
        noise = {'SearchResults': {'Parts': [{
            'ManufacturerPartNumber': 'CC0805KRX7R9BB104',
            'Manufacturer': 'YAGEO',
            'Availability': '5000 In Stock',
            'PriceBreaks': [{'Quantity': 1, 'Price': '$0.02'}],
        }]}}
        self.assertIsNone(self.client.to_record(noise, {'mpn': 'STM32F103C8T6', 'quantity': 1}))

    def test_obsolete_part_keeps_its_status_through_to_the_offer(self):
        response = {'SearchResults': {'Parts': [{
            'ManufacturerPartNumber': 'ATMEGA328P-AU',
            'Manufacturer': 'Microchip',
            'Availability': 'None',
            'LifecycleStatus': 'Obsolete',
            'SuggestedReplacement': 'ATMEGA328PB-AU',
            'LeadTime': '52 Weeks',
            'PriceBreaks': [{'Quantity': 1, 'Price': '$2.50'}],
        }]}}
        part = {'mpn': 'ATMEGA328P-AU', 'quantity': 25}
        offer = record_to_offer(self.client.to_record(response, part), part)
        self.assertEqual(offer['lifecycle'], Lifecycle.OBSOLETE)
        self.assertEqual(offer['lifecycleSeverity'], 'bad')
        self.assertEqual(offer['stock'], 0)
        self.assertFalse(offer['stockSufficient'])
        self.assertEqual(offer['suggestedReplacement'], 'ATMEGA328PB-AU')

    def test_error_payload_raises_rather_than_reading_as_empty(self):
        def boom(keyword, records=10):
            raise HttpError('Mouser API error: Invalid API key', 400, None)

        self.client.search = boom
        with self.assertRaisesRegex(HttpError, 'Invalid API key'):
            self.client.fetch_record(PART)


class FakeClient:
    configured = True

    def __init__(self, client_id, name, handler):
        self.id = client_id
        self.name = name
        self.handler = handler
        self.calls = 0

    def fetch_record(self, part):
        self.calls += 1
        return self.handler(part)


class LookupTests(unittest.TestCase):
    def setUp(self):
        self.digikey = DigiKeyClient(client_id='id', client_secret='secret')
        self.mouser = MouserClient(api_key='key')

    def both_suppliers(self):
        return [
            FakeClient('digikey', 'DigiKey', lambda p: self.digikey.to_record(DIGIKEY_RESPONSE, p)),
            FakeClient('mouser', 'Mouser', lambda p: self.mouser.to_record(MOUSER_RESPONSE, p)),
        ]

    def test_both_suppliers_pair_onto_one_row_with_a_verdict(self):
        service = LookupService(clients=self.both_suppliers(), cache=None)
        result = service.lookup_parts([
            {'row': 1, 'mpn': 'RC0603FR-0710KL', 'quantity': 500, 'manufacturer': 'Yageo'},
        ])
        rows = result['rows']
        self.assertEqual(len(rows), 1)
        self.assertEqual(len(result['suppliers']), 2)
        row = rows[0]
        self.assertEqual(row['offers']['digikey']['extendedPrice'], 7.8)
        self.assertEqual(row['offers']['mouser']['extendedPrice'], 8.5)
        self.assertEqual(row['comparison']['bestPriceSupplier'], 'DigiKey')
        # Both hold stock, so the cheapest stocked supplier is recommended.
        self.assertEqual(row['comparison']['recommendedSupplier'], 'DigiKey')
        self.assertEqual(sorted(row['comparison']['inStockSuppliers']), ['DigiKey', 'Mouser'])

    def test_repeated_part_number_costs_one_call_per_supplier(self):
        client = FakeClient('digikey', 'DigiKey', lambda p: self.digikey.to_record(DIGIKEY_RESPONSE, p))
        service = LookupService(clients=[client], cache=None)
        result = service.lookup_parts([
            {'row': 1, 'mpn': 'RC0603FR-0710KL', 'quantity': 100},
            {'row': 2, 'mpn': 'RC0603FR-0710KL', 'quantity': 100},
            {'row': 3, 'mpn': 'RC0603FR-0710KL', 'quantity': 100},
        ])
        self.assertEqual(len(result['rows']), 3)
        self.assertEqual(client.calls, 1)

    def test_cached_lookup_is_reused_on_the_next_run(self):
        cache = PartCache(ttl_seconds=60, path=None)
        client = FakeClient('digikey', 'DigiKey', lambda p: self.digikey.to_record(DIGIKEY_RESPONSE, p))
        service = LookupService(clients=[client], cache=cache)

        first = service.lookup_parts([{'mpn': 'RC0603FR-0710KL', 'quantity': 100}])
        self.assertEqual(first['stats']['apiCalls'], 1)
        self.assertEqual(first['stats']['cacheHits'], 0)

        # A different quantity reprices from the cached catalog record.
        second = service.lookup_parts([{'mpn': 'RC0603FR-0710KL', 'quantity': 5000}])
        self.assertEqual(second['stats']['apiCalls'], 0)
        self.assertEqual(second['stats']['cacheHits'], 1)
        self.assertEqual(second['rows'][0]['offers']['digikey']['supplierPartNumber'], '311-10.0KHRTR-ND')

    def test_supplier_outage_marks_that_column_errored_without_failing_the_run(self):
        def boom(part):
            raise RuntimeError('HTTP 429 from api.digikey.com')

        service = LookupService(clients=[
            FakeClient('digikey', 'DigiKey', boom),
            FakeClient('mouser', 'Mouser', lambda p: self.mouser.to_record(MOUSER_RESPONSE, p)),
        ], cache=None)

        result = service.lookup_parts([{'mpn': 'RC0603FR-0710KL', 'quantity': 100}])
        row = result['rows'][0]
        self.assertEqual(result['stats']['errors'], 1)
        self.assertFalse(row['offers']['digikey']['found'])
        self.assertTrue(row['offers']['digikey']['error'])
        self.assertIn('429', row['offers']['digikey']['reason'])
        # The surviving supplier still produces a usable answer.
        self.assertTrue(row['offers']['mouser']['found'])
        self.assertEqual(row['comparison']['recommendedSupplier'], 'Mouser')

    def test_failed_lookup_is_not_cached_so_the_next_run_retries(self):
        cache = PartCache(ttl_seconds=60, path=None)
        attempts = []

        def boom(part):
            attempts.append(1)
            raise RuntimeError('network unreachable')

        service = LookupService(clients=[FakeClient('digikey', 'DigiKey', boom)], cache=cache)
        service.lookup_parts([{'mpn': 'ABC123', 'quantity': 1}])
        service.lookup_parts([{'mpn': 'ABC123', 'quantity': 1}])
        self.assertEqual(len(attempts), 2)

    def test_progress_is_reported_once_per_supplier_lookup(self):
        service = LookupService(clients=[FakeClient('digikey', 'DigiKey', lambda p: None)], cache=None)
        seen = []
        service.lookup_parts(
            [{'mpn': 'A', 'quantity': 1}, {'mpn': 'B', 'quantity': 1}],
            on_progress=lambda p: seen.append('%d/%d' % (p['completed'], p['total'])),
        )
        self.assertEqual(sorted(seen), ['1/2', '2/2'])

    def test_running_with_no_configured_supplier_fails_loudly(self):
        service = LookupService(clients=[], cache=None)
        with self.assertRaisesRegex(RuntimeError, 'No supplier is configured'):
            service.lookup_parts([{'mpn': 'ABC', 'quantity': 1}])


class SummaryTests(unittest.TestCase):
    def test_totals_each_supplier_cart_and_the_cheapest_mix(self):
        digikey = DigiKeyClient(client_id='id', client_secret='secret')
        mouser = MouserClient(api_key='key')
        service = LookupService(clients=[
            FakeClient('digikey', 'DigiKey', lambda p: digikey.to_record(DIGIKEY_RESPONSE, p)),
            FakeClient('mouser', 'Mouser', lambda p: mouser.to_record(MOUSER_RESPONSE, p)),
        ], cache=None)

        result = service.lookup_parts([{'row': 1, 'mpn': 'RC0603FR-0710KL', 'quantity': 500}])
        summary = summarize_bom(result['rows'], result['suppliers'])

        self.assertEqual(summary['lines'], 1)
        self.assertEqual(summary['totalQuantity'], 500)
        self.assertEqual(summary['supplierTotals']['digikey']['total'], 7.8)
        self.assertEqual(summary['supplierTotals']['mouser']['total'], 8.5)
        self.assertEqual(summary['bestMixTotal'], 7.8)
        self.assertEqual(summary['cheapestSingleSource'], 'digikey')
        self.assertEqual(summary['notFoundLines'], 0)

    def test_part_nobody_carries_is_counted_as_a_risk(self):
        service = LookupService(clients=[FakeClient('digikey', 'DigiKey', lambda p: None)], cache=None)
        result = service.lookup_parts([{'mpn': 'NOSUCHPART999', 'quantity': 10}])
        summary = summarize_bom(result['rows'], result['suppliers'])
        self.assertEqual(summary['notFoundLines'], 1)
        self.assertEqual(len(summary['riskLines']), 1)
        self.assertEqual(summary['riskLines'][0]['mpn'], 'NOSUCHPART999')


if __name__ == '__main__':
    unittest.main()
