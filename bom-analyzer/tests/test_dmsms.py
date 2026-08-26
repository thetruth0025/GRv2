"""The DMSMS case form: what goes on it, and what is deliberately left blank."""

import io
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bomlib import dmsms  # noqa: E402
from bomlib.normalize import Lifecycle  # noqa: E402
from bomlib.spreadsheet import parse_xlsx  # noqa: E402


def offer(supplier, lifecycle, stock, unit=1.0, found=True):
    return {
        'supplier': supplier, 'found': found, 'lifecycle': lifecycle,
        'lifecycleSeverity': 'bad' if lifecycle == Lifecycle.OBSOLETE else 'warn',
        'stock': stock, 'totalStock': stock, 'unitPrice': unit,
        'extendedPrice': unit * 10, 'supplierPartNumber': supplier[:2] + '-1',
        'manufacturer': 'Acme',
    }


def row(mpn, lifecycle, quantity=10, stock=5000, offers=None, **extra):
    entry = {
        'index': 0, 'row': 1, 'mpn': mpn, 'quantity': quantity,
        'reference': 'R1', 'manufacturer': 'Acme', 'description': 'A part',
        'offers': offers if offers is not None else {'digikey': offer('DigiKey', lifecycle, stock)},
        'comparison': {
            'lifecycle': lifecycle,
            'lifecycleSeverity': 'bad' if lifecycle in (
                Lifecycle.OBSOLETE, Lifecycle.END_OF_LIFE, Lifecycle.DISCONTINUED) else 'warn',
            'recommendedSupplier': 'DigiKey',
            'inStockSuppliers': ['DigiKey'] if stock >= quantity else [],
            'flags': [],
        },
    }
    entry.update(extra)
    return entry


class QualifyingTests(unittest.TestCase):
    def test_every_end_of_supply_status_belongs_on_the_form(self):
        for status in (Lifecycle.OBSOLETE, Lifecycle.DISCONTINUED, Lifecycle.END_OF_LIFE,
                       Lifecycle.LAST_TIME_BUY, Lifecycle.NRND):
            self.assertTrue(dmsms.qualifies(status), status)

    def test_a_healthy_part_does_not(self):
        for status in (Lifecycle.ACTIVE, Lifecycle.NEW, Lifecycle.PREVIEW, Lifecycle.UNKNOWN):
            self.assertFalse(dmsms.qualifies(status), status)

    def test_only_the_gone_ones_are_ticked_to_begin_with(self):
        self.assertTrue(dmsms.default_selected(Lifecycle.OBSOLETE))
        self.assertTrue(dmsms.default_selected(Lifecycle.END_OF_LIFE))
        # Still buyable: listed, but the analyst decides whether it is in scope.
        self.assertFalse(dmsms.default_selected(Lifecycle.NRND))
        self.assertFalse(dmsms.default_selected(Lifecycle.LAST_TIME_BUY))

    def test_candidates_are_drawn_from_the_analysis(self):
        result = {'rows': [row('A', Lifecycle.OBSOLETE), row('B', Lifecycle.ACTIVE),
                           row('C', Lifecycle.NRND)]}
        self.assertEqual([r['mpn'] for r in dmsms.candidate_rows(result)], ['A', 'C'])


class ProvenanceTests(unittest.TestCase):
    def test_the_supplier_that_reported_the_status_is_named(self):
        entry = row('A', Lifecycle.OBSOLETE, offers={
            'digikey': offer('DigiKey', Lifecycle.ACTIVE, 100),
            'mouser': offer('Mouser', Lifecycle.OBSOLETE, 50),
        })
        # The comparison keeps the worst status, so the form must say who said it.
        self.assertEqual(dmsms.status_sources(entry), ['Mouser'])

    def test_two_suppliers_agreeing_are_both_named(self):
        entry = row('A', Lifecycle.OBSOLETE, offers={
            'digikey': offer('DigiKey', Lifecycle.OBSOLETE, 100),
            'mouser': offer('Mouser', Lifecycle.OBSOLETE, 50),
        })
        self.assertEqual(sorted(dmsms.status_sources(entry)), ['DigiKey', 'Mouser'])

    def test_stock_is_summed_across_the_suppliers_that_carry_it(self):
        entry = row('A', Lifecycle.OBSOLETE, offers={
            'digikey': offer('DigiKey', Lifecycle.OBSOLETE, 100),
            'mouser': offer('Mouser', Lifecycle.OBSOLETE, 250),
            'trustedparts': offer('TrustedParts', Lifecycle.OBSOLETE, 0, found=False),
        })
        self.assertEqual(dmsms.total_stock(entry), 350)

    def test_a_part_nobody_carries_reports_no_stock_rather_than_zero(self):
        entry = row('A', Lifecycle.OBSOLETE, offers={
            'digikey': offer('DigiKey', Lifecycle.OBSOLETE, 0, found=False),
        })
        self.assertIsNone(dmsms.total_stock(entry))


class RiskTests(unittest.TestCase):
    def test_obsolete_with_nothing_left_is_the_worst_case(self):
        self.assertEqual(dmsms.suggest_risk(row('A', Lifecycle.OBSOLETE, 100, stock=0)), 'High')

    def test_obsolete_but_still_coverable_is_not_yet_urgent(self):
        self.assertEqual(dmsms.suggest_risk(row('A', Lifecycle.OBSOLETE, 100, stock=5000)), 'Medium')

    def test_nrnd_with_stock_is_the_mildest_thing_on_the_form(self):
        self.assertEqual(dmsms.suggest_risk(row('A', Lifecycle.NRND, 10, stock=5000)), 'Low')

    def test_nrnd_you_cannot_cover_still_needs_looking_at(self):
        self.assertEqual(dmsms.suggest_risk(row('A', Lifecycle.NRND, 10, stock=2)), 'Medium')

    def test_a_last_time_buy_you_cannot_cover_is_high(self):
        self.assertEqual(dmsms.suggest_risk(row('A', Lifecycle.LAST_TIME_BUY, 100, stock=1)), 'High')


class FormTests(unittest.TestCase):
    def build(self, rows, meta=None):
        buffer = io.BytesIO()
        dmsms.write_form(buffer, rows, meta or {'program': 'Falcon II'})
        return parse_xlsx(buffer.getvalue())

    def find_case_table(self, grid):
        for index, line in enumerate(grid):
            if line and line[0] == 'Item':
                return grid[index], grid[index + 1:]
        self.fail('no case table in the form')

    def test_the_form_leads_with_the_programme_it_is_for(self):
        grid = self.build([row('A', Lifecycle.OBSOLETE)])
        self.assertEqual(grid[0][0], 'DMSMS Case Form')
        self.assertIn('Falcon II', grid[1][0])
        self.assertTrue(any(line[:2] == ['Program / platform', 'Falcon II'] for line in grid))

    def test_case_details_carry_through(self):
        grid = self.build([row('A', Lifecycle.OBSOLETE)], {
            'program': 'Osprey', 'caseNumber': 'DM-1', 'preparedBy': 'A. Buyer',
            'contract': 'N00019', 'cage': '1A2B3', 'notes': 'From the Q3 refresh.',
        })
        flat = {line[0]: line[1] for line in grid if len(line) > 1 and line[0]}
        self.assertEqual(flat['DMSMS case number'], 'DM-1')
        self.assertEqual(flat['Prepared by'], 'A. Buyer')
        self.assertEqual(flat['CAGE code'], '1A2B3')
        self.assertEqual(flat['Notes'], 'From the Q3 refresh.')

    def test_one_record_per_selected_part_in_the_order_given(self):
        grid = self.build([row('AAA', Lifecycle.OBSOLETE), row('BBB', Lifecycle.NRND)])
        header, body = self.find_case_table(grid)
        records = [line for line in body if line and str(line[0]).isdigit()]
        self.assertEqual([line[header.index('Manufacturer Part Number')] for line in records],
                         ['AAA', 'BBB'])
        self.assertEqual([line[0] for line in records], ['1', '2'])

    def test_the_analyst_columns_are_left_empty_rather_than_guessed(self):
        grid = self.build([row('AAA', Lifecycle.OBSOLETE)])
        header, body = self.find_case_table(grid)
        record = next(line for line in body if line and str(line[0]).isdigit())
        for name in dmsms.ANALYST_COLUMNS:
            self.assertEqual(record[header.index(name)], '', name)

    def test_the_assembly_a_part_sits_on_is_recorded(self):
        grid = self.build([row('AAA', Lifecycle.OBSOLETE, assembly='Main board')])
        header, body = self.find_case_table(grid)
        record = next(line for line in body if line and str(line[0]).isdigit())
        self.assertEqual(record[header.index('Next Higher Assembly')], 'Main board')

    def test_the_status_and_who_reported_it_travel_together(self):
        grid = self.build([row('AAA', Lifecycle.OBSOLETE, offers={
            'digikey': offer('DigiKey', Lifecycle.ACTIVE, 100),
            'mouser': offer('Mouser', Lifecycle.OBSOLETE, 50),
        })])
        header, body = self.find_case_table(grid)
        record = next(line for line in body if line and str(line[0]).isdigit())
        self.assertEqual(record[header.index('Lifecycle Status')], 'Obsolete')
        self.assertEqual(record[header.index('Status Source')], 'Mouser')

    def test_the_form_says_the_risk_column_is_only_a_suggestion(self):
        grid = self.build([row('AAA', Lifecycle.OBSOLETE)])
        notes = ' '.join(line[0] for line in grid if line and line[0])
        self.assertIn('not a determination', notes)
        self.assertIn('Confirm against the manufacturer', notes)

    def test_the_counts_at_the_top_match_the_records(self):
        grid = self.build([row('A', Lifecycle.OBSOLETE), row('B', Lifecycle.OBSOLETE),
                           row('C', Lifecycle.NRND)])
        flat = {line[0]: line[1] for line in grid if len(line) > 1 and line[0]}
        self.assertEqual(flat['Obsolete'], '2')
        self.assertEqual(flat['Not Recommended for New Designs'], '1')

    def test_a_part_number_that_looks_like_a_formula_stays_text(self):
        grid = self.build([row('=cmd|calc', Lifecycle.OBSOLETE)])
        header, body = self.find_case_table(grid)
        record = next(line for line in body if line and str(line[0]).isdigit())
        self.assertEqual(record[header.index('Manufacturer Part Number')], '=cmd|calc')

    def test_the_column_widths_still_line_up_with_the_columns(self):
        self.assertEqual(len(dmsms.CASE_COLUMNS), len(dmsms.CASE_WIDTHS))


if __name__ == '__main__':
    unittest.main()
