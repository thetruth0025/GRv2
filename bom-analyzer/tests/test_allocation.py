"""Covering a line from several suppliers at once.

Two rules under test. "Short" is a fact about the line, not about any one
supplier: four suppliers holding 80 each cover a need for 200 between them.
And lead time decides who supplies it, always — price only settles a field
that is already tied on speed.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bomlib.normalize import allocate_stock, compare_offers  # noqa: E402


def offer(name, stock, unit=1.0, lead=None, needed=200, moq=1, multiple=1, found=True):
    return {
        'supplier': name, 'found': found, 'stock': stock,
        'stockSufficient': None if stock is None else stock >= needed,
        'leadTimeDays': lead, 'unitPrice': unit,
        'extendedPrice': None if unit is None else round(unit * needed, 4),
        'minimumOrderQuantity': moq, 'orderMultiple': multiple,
        'priceBreaks': [{'quantity': 1, 'unitPrice': unit}] if unit is not None else [],
        'currency': 'USD', 'supplierPartNumber': name[:2].upper() + '-1',
        'lifecycle': 'Active',
    }


def plan_for(offers, quantity=200):
    return allocate_stock(offers, quantity)


class ShortIsAboutTheLineTests(unittest.TestCase):
    def test_four_suppliers_holding_eighty_each_cover_a_need_for_two_hundred(self):
        summary = compare_offers([offer(n, 80) for n in ('A', 'B', 'C', 'D')], 200)
        self.assertTrue(summary['stockCovers'])
        self.assertEqual(summary['combinedStock'], 320)
        # Not a single supplier can cover it alone, and that is not a shortage.
        self.assertEqual(summary['inStockSuppliers'], [])
        self.assertFalse(any(f['level'] == 'bad' for f in summary['flags']))

    def test_a_genuine_shortage_is_the_combined_position(self):
        summary = compare_offers([offer('A', 10), offer('B', 5)], 200)
        self.assertFalse(summary['stockCovers'])
        self.assertEqual(summary['combinedStock'], 15)
        self.assertTrue(any('Combined stock (15) is below the required 200' in f['text']
                            for f in summary['flags']))

    def test_one_supplier_holding_enough_is_not_a_split(self):
        summary = compare_offers([offer('A', 5000), offer('B', 5)], 200)
        self.assertTrue(summary['stockCovers'])
        self.assertFalse(summary['allocation']['splitRequired'])
        self.assertEqual(len(summary['allocation']['lines']), 1)

    def test_exactly_enough_between_them_counts_as_covered(self):
        summary = compare_offers([offer('A', 100), offer('B', 100)], 200)
        self.assertTrue(summary['stockCovers'])
        self.assertEqual(summary['allocation']['shortfall'], 0)

    def test_one_short_of_enough_does_not(self):
        summary = compare_offers([offer('A', 100), offer('B', 99)], 200)
        self.assertFalse(summary['stockCovers'])
        self.assertEqual(summary['allocation']['shortfall'], 1)

    def test_a_split_is_reported_as_a_note_not_as_a_risk(self):
        summary = compare_offers([offer('A', 120), offer('B', 120)], 200)
        split = [f for f in summary['flags'] if 'Split' in f['text']]
        self.assertEqual(len(split), 1)
        self.assertEqual(split[0]['level'], 'info')


class AllocationTests(unittest.TestCase):
    def takes(self, plan):
        return [(line['supplier'], line['take']) for line in plan['lines']]

    def test_the_quantity_is_drawn_until_the_need_is_met_and_no_further(self):
        plan = plan_for([offer('A', 80, unit=0.5), offer('B', 80, unit=0.6),
                         offer('C', 80, unit=0.7)])
        self.assertEqual(self.takes(plan), [('A', 80), ('B', 80), ('C', 40)])
        self.assertEqual(plan['covered'], 200)
        self.assertEqual(plan['shortfall'], 0)
        # The fourth supplier was not needed, so it is not in the plan.
        self.assertEqual(plan['suppliers'], 3)

    def test_every_draw_is_from_stock_so_the_plan_ships_today(self):
        # Each of these quotes a long factory lead time behind its shelf. The
        # units on the shelf do not wait for it.
        plan = plan_for([offer('A', 80, lead=180), offer('B', 80, lead=120),
                         offer('C', 80, lead=240)])
        self.assertEqual(plan['leadTimeDays'], 0)
        self.assertTrue(all(line['leadTimeDays'] == 0 for line in plan['lines']))

    def test_price_settles_an_order_that_is_tied_on_speed(self):
        # All three ship today, so nothing separates them but price.
        plan = plan_for([offer('Dear', 80, unit=2.0), offer('Cheap', 80, unit=0.5),
                         offer('Middle', 80, unit=1.0)])
        self.assertEqual([line['supplier'] for line in plan['lines']],
                         ['Cheap', 'Middle', 'Dear'])

    def test_a_factory_lead_time_never_outranks_stock_on_hand(self):
        plan = plan_for([offer('Slow', 500, unit=2.0, lead=180),
                         offer('Fast', 80, unit=0.5, lead=7)])
        # Both are drawn from stock, so both ship today; price orders them, and
        # the long factory quote behind Slow's shelf is beside the point.
        self.assertEqual(self.takes(plan), [('Fast', 80), ('Slow', 120)])

    def test_a_supplier_with_no_stock_is_not_in_the_plan(self):
        plan = plan_for([offer('A', 0, lead=28), offer('B', 300)])
        self.assertEqual(self.takes(plan), [('B', 200)])

    def test_the_plan_totals_what_it_would_cost(self):
        plan = plan_for([offer('A', 80, unit=0.5), offer('B', 200, unit=1.0)])
        # 80 at 0.50 plus 120 at 1.00.
        self.assertEqual(plan['total'], 160.0)
        self.assertEqual([line['extendedPrice'] for line in plan['lines']], [40.0, 120.0])

    def test_packaging_still_applies_to_a_partial_draw(self):
        # Asking for 120 from a supplier who sells in reels of 1,000 buys a reel.
        plan = plan_for([offer('A', 80, unit=0.5), offer('B', 5000, unit=1.0,
                                                         moq=1000, multiple=1000)])
        line = plan['lines'][1]
        self.assertEqual(line['take'], 120)
        self.assertEqual(line['orderQuantity'], 1000)
        self.assertEqual(line['extendedPrice'], 1000.0)

    def test_a_partial_draw_never_orders_more_than_the_supplier_holds(self):
        plan = plan_for([offer('A', 80, unit=0.5), offer('B', 150, unit=1.0,
                                                         moq=1000, multiple=1000)])
        line = plan['lines'][1]
        self.assertEqual(line['take'], 120)
        self.assertEqual(line['orderQuantity'], 150)

    def test_a_shortfall_names_who_could_factory_order_the_rest_soonest(self):
        plan = plan_for([offer('A', 10, unit=1.0, lead=180),
                         offer('B', 5, unit=2.0, lead=42)])
        self.assertEqual(plan['shortfall'], 185)
        self.assertEqual(plan['backorder'],
                         {'supplier': 'B', 'quantity': 185, 'leadTimeDays': 42})

    def test_no_shortfall_means_no_backorder_to_name(self):
        self.assertNotIn('backorder', plan_for([offer('A', 500)]))

    def test_nobody_holding_any_stock_is_an_empty_plan(self):
        plan = plan_for([offer('A', 0, lead=28), offer('B', 0, lead=56)])
        self.assertEqual(plan['lines'], [])
        self.assertEqual(plan['shortfall'], 200)
        self.assertEqual(plan['backorder']['supplier'], 'A')

    def test_a_supplier_that_carries_nothing_is_not_considered(self):
        plan = plan_for([offer('A', 80), offer('B', 500, found=False)])
        self.assertEqual(self.takes(plan), [('A', 80)])
        self.assertEqual(plan['shortfall'], 120)

    def test_an_unpriced_supplier_sorts_last_but_still_supplies(self):
        plan = plan_for([offer('Unpriced', 500, unit=None), offer('Priced', 80, unit=1.0)])
        self.assertEqual(self.takes(plan), [('Priced', 80), ('Unpriced', 120)])
        # A plan that cannot be fully priced does not pretend to a total.
        self.assertIsNone(plan['total'])


class LeadTimeReportTests(unittest.TestCase):
    """A split-covered line ships today, and the report has to say so."""

    def row(self, offers, quantity=200):
        comparison = compare_offers(offers, quantity)
        return {
            'index': 0, 'row': 1, 'mpn': 'SPLIT-1', 'quantity': quantity,
            'manufacturer': 'Acme', 'description': 'A part', 'reference': 'R1',
            'offers': {o['supplier'].lower(): o for o in offers},
            'comparison': comparison,
        }

    def summarize(self, offers, quantity=200):
        from bomlib import leadtime
        suppliers = [{'id': o['supplier'].lower(), 'name': o['supplier']} for o in offers]
        return leadtime.summarize_row(self.row(offers, quantity), suppliers)

    def test_a_split_covered_line_is_not_reported_at_a_factory_lead_time(self):
        from bomlib import leadtime
        entry = self.summarize([offer('A', 80, lead=180), offer('B', 80, lead=120),
                                offer('C', 80, lead=240)])
        self.assertEqual(entry['band'], leadtime.QUICK)
        self.assertEqual(entry['days'], 0)
        self.assertEqual(entry['availability'], 'In stock, split')

    def test_it_names_the_purchase_orders_rather_than_one_supplier(self):
        entry = self.summarize([offer('A', 80, unit=0.5), offer('B', 80, unit=1.0),
                                offer('C', 80, unit=2.0)])
        self.assertEqual(entry['supplier'], '3 suppliers, split')
        self.assertIn('Split 200 across 3 suppliers', entry['note'])
        self.assertIn('A 80, B 80, C 40', entry['note'])
        self.assertEqual(len(entry['allocation']), 3)

    def test_a_line_one_supplier_can_cover_is_unchanged(self):
        entry = self.summarize([offer('A', 5000, unit=0.5), offer('B', 80, unit=1.0)])
        self.assertEqual(entry['supplier'], 'A')
        self.assertFalse(entry['split'])

    def test_a_shortfall_is_still_reported_at_its_lead_time(self):
        from bomlib import leadtime
        entry = self.summarize([offer('A', 10, lead=120)])
        self.assertEqual(entry['band'], leadtime.LONG)
        self.assertEqual(entry['days'], 120)


if __name__ == '__main__':
    unittest.main()
