import unittest

from bomlib.normalize import (
    Lifecycle,
    build_offer,
    compare_offers,
    format_lead_time,
    missing_offer,
    normalize_lifecycle,
    order_quantity,
    parse_lead_time_days,
    parse_money,
    pick_variation,
    price_at_quantity,
    record_to_offer,
    worst_lifecycle,
)


class LifecycleTests(unittest.TestCase):
    def test_vocabulary_from_both_suppliers_maps_onto_one_scale(self):
        self.assertEqual(normalize_lifecycle('Active'), Lifecycle.ACTIVE)
        self.assertEqual(normalize_lifecycle('Obsolete'), Lifecycle.OBSOLETE)
        self.assertEqual(normalize_lifecycle('Not For New Designs'), Lifecycle.NRND)
        self.assertEqual(normalize_lifecycle('Not Recommended for New Designs'), Lifecycle.NRND)
        self.assertEqual(normalize_lifecycle('NRND'), Lifecycle.NRND)
        self.assertEqual(normalize_lifecycle('Last Time Buy'), Lifecycle.LAST_TIME_BUY)
        self.assertEqual(normalize_lifecycle('End of Life'), Lifecycle.END_OF_LIFE)
        self.assertEqual(normalize_lifecycle('Discontinued at Digi-Key'), Lifecycle.DISCONTINUED)
        self.assertEqual(normalize_lifecycle('New at Mouser'), Lifecycle.NEW)
        self.assertEqual(normalize_lifecycle(''), Lifecycle.UNKNOWN)
        self.assertEqual(normalize_lifecycle(None), Lifecycle.UNKNOWN)

    def test_worst_status_across_suppliers_wins_so_a_risk_is_never_hidden(self):
        self.assertEqual(worst_lifecycle([Lifecycle.ACTIVE, Lifecycle.OBSOLETE]), Lifecycle.OBSOLETE)
        self.assertEqual(worst_lifecycle([Lifecycle.ACTIVE, Lifecycle.NRND]), Lifecycle.NRND)
        self.assertEqual(worst_lifecycle([Lifecycle.ACTIVE, Lifecycle.ACTIVE]), Lifecycle.ACTIVE)
        self.assertEqual(worst_lifecycle([]), Lifecycle.UNKNOWN)


class LeadTimeTests(unittest.TestCase):
    def test_parses_every_unit_the_suppliers_use(self):
        self.assertEqual(parse_lead_time_days('12 Weeks'), 84)
        self.assertEqual(parse_lead_time_days('45 Days'), 45)
        self.assertEqual(parse_lead_time_days('3 Months'), 90)
        self.assertEqual(parse_lead_time_days('In Stock'), 0)
        self.assertEqual(parse_lead_time_days(16), 16)
        self.assertIsNone(parse_lead_time_days(''))
        self.assertIsNone(parse_lead_time_days(None))
        # A bare number in a lead-time field means weeks at both suppliers.
        self.assertEqual(parse_lead_time_days('20'), 140)

    def test_formats_back_into_buyer_readable_text(self):
        self.assertEqual(format_lead_time(0), 'In stock')
        self.assertEqual(format_lead_time(84), '12 weeks')
        self.assertEqual(format_lead_time(7), '1 week')
        self.assertEqual(format_lead_time(45), '45 days')
        self.assertIsNone(format_lead_time(None))


class MoneyTests(unittest.TestCase):
    def test_parses_out_of_the_strings_mouser_returns(self):
        self.assertEqual(parse_money('$1.23'), 1.23)
        self.assertEqual(parse_money('1.23'), 1.23)
        self.assertEqual(parse_money(0.0087), 0.0087)
        self.assertEqual(parse_money('$1,234.56'), 1234.56)
        # European formatting: comma is the decimal separator.
        self.assertEqual(parse_money('1.234,56 €'), 1234.56)
        self.assertIsNone(parse_money(''))


class PricingTests(unittest.TestCase):
    BREAKS = [
        {'quantity': 1, 'unitPrice': 1.0},
        {'quantity': 10, 'unitPrice': 0.8},
        {'quantity': 100, 'unitPrice': 0.5},
    ]

    def test_applicable_break_is_highest_at_or_below_quantity(self):
        self.assertEqual(price_at_quantity(self.BREAKS, 1)['unitPrice'], 1.0)
        self.assertEqual(price_at_quantity(self.BREAKS, 9)['unitPrice'], 1.0)
        self.assertEqual(price_at_quantity(self.BREAKS, 10)['unitPrice'], 0.8)
        self.assertEqual(price_at_quantity(self.BREAKS, 250)['unitPrice'], 0.5)
        self.assertIsNone(price_at_quantity([], 10))

    def test_below_smallest_break_the_buyer_still_pays_that_break(self):
        self.assertEqual(price_at_quantity([{'quantity': 10, 'unitPrice': 0.8}], 1)['unitPrice'], 0.8)

    def test_order_quantity_respects_minimum_and_multiples(self):
        self.assertEqual(order_quantity(100, 1, 1), 100)
        self.assertEqual(order_quantity(5, 10, 1), 10)
        self.assertEqual(order_quantity(105, 1, 50), 150)
        self.assertEqual(order_quantity(3, 1, 1), 3)
        self.assertEqual(order_quantity(0, 1, 1), 1)

    def test_offer_prices_the_quantity_actually_purchased(self):
        offer = build_offer({
            'supplier': 'Mouser',
            'quantity': 105,
            'minimumOrderQuantity': 1,
            'orderMultiple': 50,
            'priceBreaks': [
                {'quantity': 1, 'unitPrice': 0.1},
                {'quantity': 100, 'unitPrice': 0.05},
            ],
            'stock': 5000,
            'leadTime': '8 Weeks',
            'lifecycle': 'Active',
        })
        self.assertEqual(offer['orderQuantity'], 150)
        self.assertEqual(offer['unitPrice'], 0.05)
        self.assertEqual(offer['extendedPrice'], 7.5)
        self.assertTrue(offer['stockSufficient'])
        self.assertEqual(offer['leadTimeDays'], 56)
        self.assertEqual(offer['lifecycle'], Lifecycle.ACTIVE)

    def test_stock_below_order_quantity_is_insufficient(self):
        offer = build_offer({
            'supplier': 'DigiKey',
            'quantity': 500,
            'stock': 100,
            'priceBreaks': [{'quantity': 1, 'unitPrice': 1}],
        })
        self.assertFalse(offer['stockSufficient'])


class PackagingTests(unittest.TestCase):
    def test_prefers_an_option_that_covers_the_quantity_then_price(self):
        variations = [
            {'supplierPartNumber': 'CUT-TAPE', 'stock': 50, 'minimumOrderQuantity': 1,
             'orderMultiple': 1, 'priceBreaks': [{'quantity': 1, 'unitPrice': 0.1}]},
            {'supplierPartNumber': 'REEL', 'stock': 10000, 'minimumOrderQuantity': 1,
             'orderMultiple': 1, 'priceBreaks': [{'quantity': 1, 'unitPrice': 0.12}]},
        ]
        self.assertEqual(pick_variation(variations, 50)['supplierPartNumber'], 'CUT-TAPE')
        self.assertEqual(pick_variation(variations, 500)['supplierPartNumber'], 'REEL')

    def test_low_unit_price_loses_when_its_minimum_forces_a_larger_buy(self):
        variations = [
            {'supplierPartNumber': 'CUT-TAPE', 'stock': 5000, 'minimumOrderQuantity': 1,
             'orderMultiple': 1, 'priceBreaks': [{'quantity': 100, 'unitPrice': 0.0156}]},
            {'supplierPartNumber': 'REEL', 'stock': 1000000, 'minimumOrderQuantity': 5000,
             'orderMultiple': 5000, 'priceBreaks': [{'quantity': 5000, 'unitPrice': 0.00518}]},
        ]
        # The reel's unit price is a third of cut tape, but buying 500 pieces
        # costs $7.80 on cut tape against $25.90 for a minimum reel.
        self.assertEqual(pick_variation(variations, 500)['supplierPartNumber'], 'CUT-TAPE')
        # Past the cut-tape stock the reel is the only option that can ship.
        self.assertEqual(pick_variation(variations, 20000)['supplierPartNumber'], 'REEL')

    def test_marketplace_is_a_last_resort_even_when_cheapest(self):
        variations = [
            {'supplierPartNumber': 'MARKET', 'stock': 10000, 'minimumOrderQuantity': 1,
             'orderMultiple': 1, 'priceBreaks': [{'quantity': 1, 'unitPrice': 0.05}], 'marketPlace': True},
            {'supplierPartNumber': 'STOCK', 'stock': 10000, 'minimumOrderQuantity': 1,
             'orderMultiple': 1, 'priceBreaks': [{'quantity': 1, 'unitPrice': 0.09}], 'marketPlace': False},
        ]
        self.assertEqual(pick_variation(variations, 100)['supplierPartNumber'], 'STOCK')

    def test_cached_record_reprices_at_a_different_quantity(self):
        record = {
            'supplier': 'DigiKey',
            'manufacturerPartNumber': 'RC0603FR-0710KL',
            'leadTime': '10 Weeks',
            'lifecycle': 'Active',
            'totalStock': 100000,
            'currency': 'USD',
            'variations': [{
                'supplierPartNumber': '311-10.0KHRCT-ND',
                'stock': 100000,
                'minimumOrderQuantity': 1,
                'orderMultiple': 1,
                'priceBreaks': [
                    {'quantity': 1, 'unitPrice': 0.1},
                    {'quantity': 1000, 'unitPrice': 0.01},
                ],
            }],
        }
        self.assertEqual(record_to_offer(record, {'quantity': 10})['unitPrice'], 0.1)
        self.assertEqual(record_to_offer(record, {'quantity': 5000})['unitPrice'], 0.01)
        self.assertEqual(record_to_offer(record, {'quantity': 5000})['extendedPrice'], 50)


def offer(supplier, **overrides):
    spec = {
        'supplier': supplier,
        'quantity': 100,
        'stock': 1000,
        'leadTime': '10 Weeks',
        'lifecycle': 'Active',
        'priceBreaks': [{'quantity': 1, 'unitPrice': 1}],
    }
    spec.update(overrides)
    return build_offer(spec)


class ComparisonTests(unittest.TestCase):
    def test_cheaper_supplier_wins_on_price_and_is_named(self):
        summary = compare_offers([
            offer('DigiKey', priceBreaks=[{'quantity': 1, 'unitPrice': 1}]),
            offer('Mouser', priceBreaks=[{'quantity': 1, 'unitPrice': 0.75}]),
        ], 100)
        self.assertEqual(summary['bestPriceSupplier'], 'Mouser')
        self.assertEqual(summary['bestPrice'], 75)
        self.assertEqual(summary['priceSpread'], 25)

    def test_stock_on_hand_beats_a_shorter_quoted_factory_lead_time(self):
        summary = compare_offers([
            offer('DigiKey', stock=0, leadTime='2 Weeks'),
            offer('Mouser', stock=5000, leadTime='30 Weeks'),
        ], 100)
        self.assertEqual(summary['inStockSuppliers'], ['Mouser'])
        self.assertEqual(summary['bestLeadTimeDays'], 0)
        self.assertEqual(summary['bestLeadTimeSupplier'], 'Mouser')
        self.assertEqual(summary['recommendedSupplier'], 'Mouser')

    def test_with_both_in_stock_no_lead_time_winner_is_invented(self):
        summary = compare_offers([
            offer('DigiKey', stock=5000, leadTime='30 Weeks', priceBreaks=[{'quantity': 1, 'unitPrice': 1}]),
            offer('Mouser', stock=5000, leadTime='2 Weeks', priceBreaks=[{'quantity': 1, 'unitPrice': 2}]),
        ], 100)
        # Both ship today, so neither factory lead time is the one to act on.
        self.assertEqual(summary['bestLeadTimeDays'], 0)
        self.assertIsNone(summary['bestLeadTimeSupplier'])
        self.assertEqual(sorted(summary['bestLeadTimeSuppliers']), ['DigiKey', 'Mouser'])
        # The cheaper of the two equally fast suppliers is the recommendation.
        self.assertEqual(summary['recommendedSupplier'], 'DigiKey')

    def test_with_nobody_holding_stock_the_shortest_lead_time_is_recommended(self):
        summary = compare_offers([
            offer('DigiKey', stock=0, leadTime='20 Weeks'),
            offer('Mouser', stock=0, leadTime='6 Weeks'),
        ], 100)
        self.assertEqual(summary['recommendedSupplier'], 'Mouser')
        self.assertEqual(summary['bestLeadTimeDays'], 42)

    def test_a_split_names_who_supplies_how_many(self):
        summary = compare_offers([offer('DigiKey', stock=60), offer('Mouser', stock=60)], 100)
        texts = [f['text'] for f in summary['flags']]
        self.assertTrue(any('Split 100 across 2 suppliers' in t for t in texts), texts)
        self.assertTrue(any('DigiKey 60, Mouser 40' in t for t in texts), texts)
        # Coverable is not short, however many purchase orders it takes.
        self.assertTrue(summary['stockCovers'])
        self.assertFalse(any(f['level'] == 'bad' for f in summary['flags']), texts)

    def test_no_stock_but_a_quoted_date_is_a_factory_order_not_an_emergency(self):
        summary = compare_offers([offer('DigiKey', stock=10, leadTime='10 Weeks'),
                                  offer('Mouser', stock=5, leadTime='20 Weeks')], 100)
        texts = [f['text'] for f in summary['flags']]
        self.assertFalse(summary['stockCovers'])
        self.assertTrue(summary['obtainable'])
        self.assertTrue(any('No stock on hand — soonest is 10 weeks' in t for t in texts), texts)
        self.assertFalse(any(f['level'] == 'bad' for f in summary['flags']), texts)

    def test_short_with_nobody_quoting_a_date_is_the_emergency(self):
        summary = compare_offers([offer('DigiKey', stock=10, leadTime=None),
                                  offer('Mouser', stock=5, leadTime=None)], 100)
        texts = [f['text'] for f in summary['flags']]
        self.assertFalse(summary['obtainable'])
        self.assertTrue(any('below the required 100' in t and 'no supplier quoted a date' in t
                            for t in texts), texts)
        self.assertTrue(any(f['level'] == 'bad' for f in summary['flags']), texts)

    def test_obsolete_from_one_supplier_marks_the_whole_line(self):
        summary = compare_offers([
            offer('DigiKey', lifecycle='Active'),
            offer('Mouser', lifecycle='Obsolete'),
        ], 100)
        self.assertEqual(summary['lifecycle'], Lifecycle.OBSOLETE)
        self.assertEqual(summary['lifecycleSeverity'], 'bad')
        self.assertTrue(any('find a replacement' in f['text'] for f in summary['flags']))

    def test_part_no_supplier_carries_is_flagged_not_silently_priced(self):
        summary = compare_offers([
            missing_offer('DigiKey', 'no match'),
            missing_offer('Mouser', 'no match'),
        ], 100)
        self.assertIsNone(summary['recommendedSupplier'])
        self.assertIsNone(summary['bestPrice'])
        self.assertTrue(any('Not found at any' in f['text'] for f in summary['flags']))

    def test_single_source_part_is_called_out(self):
        summary = compare_offers([offer('DigiKey'), missing_offer('Mouser', 'no match')], 100)
        self.assertTrue(any('Single source' in f['text'] for f in summary['flags']))

    def test_long_lead_time_is_surfaced_even_when_all_else_is_fine(self):
        summary = compare_offers([
            offer('DigiKey', stock=0, leadTime='40 Weeks'),
            offer('Mouser', stock=0, leadTime='52 Weeks'),
        ], 100)
        self.assertTrue(any('Best lead time is' in f['text'] for f in summary['flags']))


if __name__ == '__main__':
    unittest.main()
