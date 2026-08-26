import argparse
import io
import json
import os
import sys
import tempfile
import unittest

import bom
from bomlib.lookup import LookupService, summarize_bom
from bomlib.report import build_header, build_rows, write_csv, write_json, write_workbook
from bomlib.spreadsheet import parse_delimited, parse_xlsx


class StubSupplier:
    """A supplier that always answers, so CLI behaviour can be tested offline."""

    configured = True

    def __init__(self, client_id, name, unit_price, stock=10000, lifecycle='Active',
                 lead='10 Weeks', found=True):
        self.id = client_id
        self.name = name
        self.unit_price = unit_price
        self.stock = stock
        self.lifecycle = lifecycle
        self.lead = lead
        self.found = found

    def fetch_record(self, part):
        if not self.found:
            return None
        return {
            'supplier': self.name,
            'manufacturer': 'Acme',
            'manufacturerPartNumber': part['mpn'],
            'description': 'Test part ' + part['mpn'],
            'leadTime': self.lead,
            'lifecycle': self.lifecycle,
            'totalStock': self.stock,
            'currency': 'USD',
            'variations': [{
                'supplierPartNumber': self.id[:2].upper() + '-' + part['mpn'],
                'stock': self.stock,
                'minimumOrderQuantity': 1,
                'orderMultiple': 1,
                'priceBreaks': [{'quantity': 1, 'unitPrice': self.unit_price}],
            }],
        }


def analyze(parts, clients=None):
    clients = clients or [
        StubSupplier('digikey', 'DigiKey', 0.10),
        StubSupplier('mouser', 'Mouser', 0.20),
    ]
    service = LookupService(clients=clients, cache=None)
    result = service.lookup_parts(parts)
    return result, summarize_bom(result['rows'], result['suppliers'])


class PastedInputTests(unittest.TestCase):
    def test_reads_one_part_per_line_with_optional_quantity(self):
        lines = bom.parse_pasted('ABC123, 100\nDEF456\t25\nGHI789\n')
        self.assertEqual([l['mpn'] for l in lines], ['ABC123', 'DEF456', 'GHI789'])
        self.assertEqual([l['quantity'] for l in lines], [100, 25, 1])

    def test_blank_lines_and_comments_are_ignored(self):
        lines = bom.parse_pasted('\n# a note\nABC123, 5\n\n')
        self.assertEqual(len(lines), 1)
        self.assertEqual(lines[0]['mpn'], 'ABC123')

    def test_a_third_field_is_taken_as_the_manufacturer(self):
        lines = bom.parse_pasted('ABC123, 5, Murata')
        self.assertEqual(lines[0]['manufacturer'], 'Murata')

    def test_an_unparseable_quantity_falls_back_to_one(self):
        self.assertEqual(bom.parse_pasted('ABC123, lots')[0]['quantity'], 1)


class ColumnOverrideTests(unittest.TestCase):
    HEADERS = ['Line', 'Widget Code', 'Count']

    def override(self, **kwargs):
        args = argparse.Namespace()
        for field in ('mpn', 'quantity', 'reference', 'manufacturer', 'description', 'footprint'):
            setattr(args, '%s_column' % field, kwargs.get(field))
        return bom.apply_overrides({'mpn': 0}, self.HEADERS, args)

    def test_a_column_can_be_named(self):
        self.assertEqual(self.override(mpn='Widget Code')['mpn'], 1)

    def test_naming_is_case_insensitive(self):
        self.assertEqual(self.override(mpn='widget code')['mpn'], 1)

    def test_a_column_can_be_given_by_index(self):
        self.assertEqual(self.override(mpn='2')['mpn'], 2)

    def test_an_empty_override_clears_the_field(self):
        self.assertNotIn('mpn', self.override(mpn=''))

    def test_an_unknown_name_is_refused_with_the_available_columns(self):
        with self.assertRaises(ValueError) as caught:
            self.override(mpn='Nope')
        self.assertIn('Widget Code', str(caught.exception))

    def test_an_out_of_range_index_is_refused(self):
        with self.assertRaises(ValueError) as caught:
            self.override(mpn='99')
        self.assertIn('out of range', str(caught.exception))

    def test_fields_without_an_override_are_left_alone(self):
        self.assertEqual(self.override(quantity='2'), {'mpn': 0, 'quantity': 2})


class ExitCodeTests(unittest.TestCase):
    def test_never_always_succeeds(self):
        self.assertEqual(bom.check_exit('never', {'notFoundLines': 3, 'riskLines': [1]}), 0)

    def test_notfound_trips_only_on_missing_parts(self):
        self.assertEqual(bom.check_exit('notfound', {'notFoundLines': 1, 'riskLines': []}), 2)
        self.assertEqual(bom.check_exit('notfound', {'notFoundLines': 0, 'riskLines': [1]}), 0)

    def test_risk_trips_on_any_flagged_line(self):
        self.assertEqual(bom.check_exit('risk', {'notFoundLines': 0, 'riskLines': [1]}), 2)
        self.assertEqual(bom.check_exit('risk', {'notFoundLines': 0, 'riskLines': []}), 0)


class ReportTests(unittest.TestCase):
    def setUp(self):
        self.result, self.summary = analyze([
            {'row': 2, 'mpn': 'ABC123', 'quantity': 100, 'reference': 'R1'},
            {'row': 3, 'mpn': 'DEF456', 'quantity': 50, 'reference': 'C1'},
        ])

    def test_the_header_names_every_supplier_column_and_the_currency(self):
        header = build_header(self.result['suppliers'], 'USD')
        self.assertEqual(header[:3], ['Row', 'Part Number', 'Quantity'])
        self.assertIn('DigiKey Extended Price (USD)', header)
        self.assertIn('Mouser Lifecycle', header)
        self.assertEqual(header[-1], 'Notes')

    def test_every_row_matches_the_header_width(self):
        rows = build_rows(self.result, self.summary)
        width = len(rows[0])
        self.assertEqual(width, len(build_header(self.result['suppliers'], 'USD')))
        for row in rows[1:]:
            self.assertEqual(len(row), width)

    def test_a_missing_supplier_leaves_aligned_blanks_and_a_reason(self):
        result, summary = analyze(
            [{'row': 2, 'mpn': 'ABC123', 'quantity': 10}],
            clients=[
                StubSupplier('digikey', 'DigiKey', 0.10),
                StubSupplier('mouser', 'Mouser', 0.20, found=False),
            ],
        )
        rows = build_rows(result, summary)
        self.assertEqual(len(rows[1]), len(rows[0]))
        self.assertIn('No Mouser match', rows[1][-7])


class WriterTests(unittest.TestCase):
    def setUp(self):
        self.result, self.summary = analyze([
            {'row': 2, 'mpn': 'ABC123', 'quantity': 100, 'reference': 'R1'},
        ])
        self.paths = []

    def tearDown(self):
        for path in self.paths:
            if os.path.exists(path):
                os.unlink(path)

    def temp(self, suffix):
        handle, path = tempfile.mkstemp(suffix=suffix)
        os.close(handle)
        self.paths.append(path)
        return path

    def test_csv_round_trips_with_the_expected_values(self):
        path = self.temp('.csv')
        write_csv(path, self.result, self.summary)
        with open(path, encoding='utf-8-sig') as handle:
            grid = parse_delimited(handle.read())
        self.assertEqual(grid[0][:3], ['Row', 'Part Number', 'Quantity'])
        self.assertEqual(grid[1][1], 'ABC123')
        self.assertIn('10.0', grid[1])  # DigiKey extended: 100 × $0.10

    def test_csv_neutralises_values_a_spreadsheet_would_treat_as_formulas(self):
        result, summary = analyze([{'row': 2, 'mpn': '=cmd|calc', 'quantity': 1}])
        path = self.temp('.csv')
        write_csv(path, result, summary)
        with open(path, encoding='utf-8-sig') as handle:
            body = handle.read()
        self.assertIn("'=cmd|calc", body)

    def test_a_negative_number_is_not_mistaken_for_a_formula(self):
        from bomlib.report import _csv_cell
        self.assertEqual(_csv_cell(-12.5), '-12.5')
        self.assertEqual(_csv_cell('-12.5'), '-12.5')
        self.assertEqual(_csv_cell('-lead'), "'-lead")

    def test_the_workbook_leads_with_the_report_sheet(self):
        path = self.temp('.xlsx')
        write_workbook(path, self.result, self.summary, {'name': 'Widget board'})
        with open(path, 'rb') as handle:
            data = handle.read()
        # The sheet somebody opens first is the readable one, not the audit
        # trail: title, then the headline numbers.
        grid = parse_xlsx(data)
        self.assertEqual(grid[0][0], 'BOM Supplier Report')
        self.assertIn('Widget board', grid[1][0])
        self.assertEqual(grid[3][0], 'Overview')
        self.assertEqual(grid[4][:2], ['Lines', 'Units'])

    def test_the_workbook_opens_and_keeps_numbers_numeric(self):
        path = self.temp('.xlsx')
        write_workbook(path, self.result, self.summary)
        with open(path, 'rb') as handle:
            data = handle.read()
        grid = parse_xlsx(data, 'Full comparison')
        self.assertEqual(grid[0][1], 'Part Number')
        self.assertEqual(grid[1][1], 'ABC123')
        self.assertEqual(grid[1][2], '100')

    def test_the_parts_sheet_holds_one_actionable_line_per_part(self):
        path = self.temp('.xlsx')
        write_workbook(path, self.result, self.summary)
        with open(path, 'rb') as handle:
            grid = parse_xlsx(handle.read(), 'Parts')
        self.assertEqual(grid[0][:3], ['Row', 'Part Number', 'Qty'])
        self.assertEqual(grid[1][1], 'ABC123')

    def test_skipped_lines_get_their_own_sheet_only_when_there_are_some(self):
        path = self.temp('.xlsx')
        write_workbook(path, self.result, self.summary)
        with open(path, 'rb') as handle:
            self.assertRaises(ValueError, parse_xlsx, handle.read(), 'Skipped')

        result = dict(self.result)
        result['excluded'] = [{
            'row': 7, 'mpn': 'ASY0-9001', 'quantity': 2, 'reference': None,
            'description': 'Top level assembly', 'reason': 'ignored',
            'detail': 'In-house part number (starts with ASY0)',
        }]
        path = self.temp('.xlsx')
        write_workbook(path, result, self.summary)
        with open(path, 'rb') as handle:
            grid = parse_xlsx(handle.read(), 'Skipped')
        self.assertEqual(grid[1][1], 'ASY0-9001')
        self.assertIn('ASY0', grid[1][5])

    def test_json_carries_the_full_result_for_scripting(self):
        path = self.temp('.json')
        write_json(path, self.result, self.summary)
        with open(path, encoding='utf-8') as handle:
            data = json.load(handle)
        self.assertEqual(sorted(data), ['excluded', 'rows', 'stats', 'summary', 'suppliers'])
        self.assertEqual(len(data['rows']), 1)
        self.assertEqual(data['summary']['bestMixTotal'], 10.0)


class EndToEndTests(unittest.TestCase):
    """Drive main() the way a user would, with the suppliers stubbed out."""

    def setUp(self):
        self.paths = []
        self._original = bom.build_service
        bom.build_service = self._stub_service

    def tearDown(self):
        bom.build_service = self._original
        for path in self.paths:
            if os.path.exists(path):
                os.unlink(path)

    def _stub_service(self, args):
        clients = [
            StubSupplier('digikey', 'DigiKey', 0.10),
            StubSupplier('mouser', 'Mouser', 0.20, lifecycle='Obsolete'),
        ]
        if args.supplier:
            clients = [c for c in clients if c.id in set(args.supplier)]
        return LookupService(clients=clients, cache=None), None

    def temp(self, suffix, content=None):
        handle, path = tempfile.mkstemp(suffix=suffix)
        os.close(handle)
        self.paths.append(path)
        if content is not None:
            with open(path, 'w', encoding='utf-8') as out:
                out.write(content)
        return path

    def bom_file(self):
        return self.temp('.csv', 'Reference,Qty,Manufacturer Part Number\nR1,100,ABC123\nC1,50,DEF456\n')

    def test_a_run_writes_a_workbook_and_succeeds(self):
        source = self.bom_file()
        out = self.temp('.xlsx')
        self.assertEqual(bom.main([source, '-o', out, '--quiet', '--no-color']), 0)
        with open(out, 'rb') as handle:
            grid = parse_xlsx(handle.read(), 'Full comparison')
        self.assertEqual(len(grid), 3)
        self.assertEqual(grid[1][1], 'ABC123')

    def test_build_quantity_multiplies_every_line(self):
        source = self.bom_file()
        out = self.temp('.json')
        self.assertEqual(bom.main([source, '-o', out, '-b', '10', '--quiet', '--no-color']), 0)
        with open(out, encoding='utf-8') as handle:
            data = json.load(handle)
        self.assertEqual([r['quantity'] for r in data['rows']], [1000, 500])
        self.assertEqual(data['summary']['totalQuantity'], 1500)

    def test_supplier_selection_produces_a_single_column_group(self):
        source = self.bom_file()
        out = self.temp('.json')
        self.assertEqual(bom.main([source, '-o', out, '-s', 'mouser', '--quiet', '--no-color']), 0)
        with open(out, encoding='utf-8') as handle:
            data = json.load(handle)
        self.assertEqual([s['id'] for s in data['suppliers']], ['mouser'])

    def test_limit_truncates_the_bom(self):
        source = self.bom_file()
        out = self.temp('.json')
        self.assertEqual(bom.main([source, '-o', out, '--limit', '1', '--quiet', '--no-color']), 0)
        with open(out, encoding='utf-8') as handle:
            self.assertEqual(len(json.load(handle)['rows']), 1)

    def test_fail_on_risk_reports_a_nonzero_exit(self):
        source = self.bom_file()
        # The Mouser stub reports every part obsolete, so every line is flagged.
        self.assertEqual(bom.main([source, '--fail-on', 'risk', '--quiet', '--no-color']), 2)
        self.assertEqual(bom.main([source, '--quiet', '--no-color']), 0)

    def test_a_bom_with_no_part_numbers_is_an_error(self):
        source = self.temp('.csv', 'alpha,beta\n1,2\n')
        self.assertEqual(bom.main([source, '--quiet', '--no-color']), 1)

    def test_a_missing_file_is_an_error(self):
        self.assertEqual(bom.main(['/nonexistent/bom.csv', '--quiet', '--no-color']), 1)

    def test_an_unguessable_output_extension_is_an_error(self):
        source = self.bom_file()
        self.assertEqual(bom.main([source, '-o', '/tmp/x.pdf', '--quiet', '--no-color']), 1)

    def test_an_explicit_format_overrides_the_extension(self):
        source = self.bom_file()
        out = self.temp('.dat')
        self.assertEqual(bom.main([source, '-o', out, '-f', 'json', '--quiet', '--no-color']), 0)
        with open(out, encoding='utf-8') as handle:
            self.assertIn('rows', json.load(handle))

    def test_a_column_override_rescues_an_unrecognizable_header(self):
        source = self.temp('.csv', 'Line,Widget Code,Count\n1,ABC123,25\n')
        out = self.temp('.json')
        self.assertEqual(bom.main([
            source, '-o', out, '--mpn-column', 'Widget Code',
            '--quantity-column', 'Count', '--quiet', '--no-color',
        ]), 0)
        with open(out, encoding='utf-8') as handle:
            data = json.load(handle)
        self.assertEqual(data['rows'][0]['mpn'], 'ABC123')
        self.assertEqual(data['rows'][0]['quantity'], 25)

    def test_a_bad_column_override_is_an_error(self):
        source = self.bom_file()
        self.assertEqual(bom.main([source, '--mpn-column', 'Nope', '--quiet', '--no-color']), 1)

    def test_build_quantity_below_one_is_refused(self):
        source = self.bom_file()
        # Zero is falsy, so it has to be checked explicitly rather than by
        # truthiness, or it silently means "leave quantities alone".
        self.assertEqual(bom.main([source, '-b', '0', '--quiet', '--no-color']), 1)
        self.assertEqual(bom.main([source, '-b', '-5', '--quiet', '--no-color']), 1)

    def test_limit_below_one_is_refused(self):
        source = self.bom_file()
        self.assertEqual(bom.main([source, '--limit', '0', '--quiet', '--no-color']), 1)


if __name__ == '__main__':
    unittest.main()


class NameJoiningTests(unittest.TestCase):
    def test_supplier_names_read_as_a_sentence(self):
        self.assertEqual(bom._join_names([]), '')
        self.assertEqual(bom._join_names(['DigiKey']), 'DigiKey')
        self.assertEqual(bom._join_names(['DigiKey', 'Mouser']), 'DigiKey and Mouser')
        self.assertEqual(
            bom._join_names(['DigiKey', 'Mouser', 'TrustedParts']),
            'DigiKey, Mouser and TrustedParts',
        )


class ScreeningCliTests(unittest.TestCase):
    """The CLI screens the same lines the web app does."""

    def setUp(self):
        self.paths = []
        self._original = bom.build_service
        bom.build_service = EndToEndTests._stub_service.__get__(self, EndToEndTests)

    def tearDown(self):
        bom.build_service = self._original
        for path in self.paths:
            if os.path.exists(path):
                os.unlink(path)

    def temp(self, suffix):
        handle, path = tempfile.mkstemp(suffix=suffix)
        os.close(handle)
        self.paths.append(path)
        return path

    def source(self, rows):
        path = self.temp('.csv')
        with open(path, 'w', encoding='utf-8') as handle:
            handle.write('Manufacturer Part Number,Qty\n')
            for mpn, qty in rows:
                handle.write('%s,%s\n' % (mpn, qty))
        return path

    def run_to_json(self, rows, extra=None):
        out = self.temp('.json')
        argv = [self.source(rows), '-o', out, '--quiet', '--no-color'] + (extra or [])
        self.assertEqual(bom.main(argv), 0)
        with open(out, encoding='utf-8') as handle:
            return json.load(handle)

    def test_in_house_part_numbers_never_reach_a_supplier(self):
        data = self.run_to_json([('ASY0-1', 1), ('ABC123', 10), ('PCB0-7', 1)])
        self.assertEqual([r['mpn'] for r in data['rows']], ['ABC123'])
        self.assertEqual(sorted(e['mpn'] for e in data['excluded']), ['ASY0-1', 'PCB0-7'])

    def test_repeated_lines_are_merged_into_one_purchase(self):
        data = self.run_to_json([('ABC123', 10), ('ABC123', 25)])
        self.assertEqual([r['quantity'] for r in data['rows']], [35])
        self.assertEqual(data['excluded'][0]['reason'], 'merged')

    def test_screening_can_be_turned_off(self):
        data = self.run_to_json([('ASY0-1', 1), ('ABC123', 1)],
                                ['--no-ignore-prefixes'])
        self.assertEqual(sorted(r['mpn'] for r in data['rows']), ['ABC123', 'ASY0-1'])

    def test_a_custom_prefix_list_replaces_the_default(self):
        data = self.run_to_json([('ASY0-1', 1), ('FIX0-2', 1)],
                                ['--ignore-prefix', 'fix0'])
        self.assertEqual(sorted(r['mpn'] for r in data['rows']), ['ASY0-1'])

    def test_merging_happens_before_the_build_multiplier(self):
        data = self.run_to_json([('ABC123', 10), ('ABC123', 5)], ['-b', '10'])
        self.assertEqual([r['quantity'] for r in data['rows']], [150])

    def test_a_part_can_be_looked_up_without_a_bom_file(self):
        out = self.temp('.json')
        self.assertEqual(bom.main(['--part', 'ABC123,25', '--part', 'DEF456',
                                   '-o', out, '--quiet', '--no-color']), 0)
        with open(out, encoding='utf-8') as handle:
            data = json.load(handle)
        self.assertEqual([r['mpn'] for r in data['rows']], ['ABC123', 'DEF456'])
        self.assertEqual([r['quantity'] for r in data['rows']], [25, 1])

    def test_naming_a_part_directly_is_not_screened_out_as_in_house(self):
        out = self.temp('.json')
        self.assertEqual(bom.main(['--part', 'ASY0-9001', '-o', out,
                                   '--quiet', '--no-color']), 0)
        with open(out, encoding='utf-8') as handle:
            self.assertEqual([r['mpn'] for r in json.load(handle)['rows']], ['ASY0-9001'])

    def test_an_explicit_prefix_list_still_applies_to_named_parts(self):
        out = self.temp('.json')
        errors = io.StringIO()
        original = sys.stderr
        sys.stderr = errors
        try:
            code = bom.main(['--part', 'ASY0-9001', '--ignore-prefix', 'ASY0',
                             '-o', out, '--quiet', '--no-color'])
        finally:
            sys.stderr = original
        self.assertEqual(code, 1)
        self.assertIn('skipped', errors.getvalue())

    def test_no_source_at_all_says_what_the_options_are(self):
        errors = io.StringIO()
        original = sys.stderr
        sys.stderr = errors
        try:
            code = bom.main(['--quiet', '--no-color'])
        finally:
            sys.stderr = original
        self.assertEqual(code, 1)
        self.assertIn('--part', errors.getvalue())

    def test_a_bom_of_nothing_but_in_house_numbers_fails_with_advice(self):
        source = self.source([('ASY0-1', 1), ('CBL0-2', 1)])
        errors = io.StringIO()
        original = sys.stderr
        sys.stderr = errors
        try:
            code = bom.main([source, '--quiet', '--no-color'])
        finally:
            sys.stderr = original
        self.assertEqual(code, 1)
        self.assertIn('--no-ignore-prefixes', errors.getvalue())


class ReportSheetTests(unittest.TestCase):
    """The concise sheets a buyer reads, rather than the audit trail."""

    def setUp(self):
        self.result, self.summary = analyze([
            {'row': 1, 'mpn': 'ABC123', 'quantity': 100, 'reference': 'R1',
             'manufacturer': 'Acme', 'description': '10k'},
        ])

    def test_the_report_leads_with_a_title_and_the_headline_numbers(self):
        from bomlib.report import build_report_rows
        rows = build_report_rows(self.result, self.summary, {'name': 'Widget', 'generated': 'now'})
        flat = [[getattr(c, 'value', c) for c in row] for row in rows]
        self.assertEqual(flat[0][0], 'BOM Supplier Report')
        self.assertEqual(flat[1][0], 'Widget · now · prices in USD')
        self.assertIn(['Overview'] + [''] * 7, flat)
        self.assertEqual(flat[4][:3], ['Lines', 'Units', 'Best-mix total'])
        self.assertEqual(flat[5][0], 1)

    def test_a_clean_bom_says_so_instead_of_showing_an_empty_table(self):
        from bomlib.report import build_report_rows
        flat = [[getattr(c, 'value', c) for c in row]
                for row in build_report_rows(self.result, self.summary)]
        self.assertTrue(any('Needs a decision (0)' in str(row[0]) for row in flat))
        self.assertTrue(any('in production' in str(row[0]) for row in flat))

    def test_the_parts_sheet_prices_the_recommended_supplier(self):
        from bomlib.report import build_parts_rows
        rows = build_parts_rows(self.result, self.summary)
        header, first = rows[0], rows[1]
        self.assertNotIn('Also In', header)
        self.assertEqual(first[header.index('Buy From')],
                         self.result['rows'][0]['comparison']['recommendedSupplier'])
        self.assertEqual(first[header.index('Qty')], 100)

    def test_cross_bom_demand_only_adds_a_column_when_there_is_some(self):
        from bomlib.report import build_parts_rows
        self.result['rows'][0]['alsoIn'] = [{'name': 'Board B', 'quantity': 250}]
        rows = build_parts_rows(self.result, self.summary)
        header, first = rows[0], rows[1]
        self.assertIn('Also In', header)
        self.assertEqual(first[header.index('Also In')], 'Board B (250)')
        # The extra column must not shift the ones after it.
        self.assertEqual(first[header.index('Qty')], 100)
        self.assertEqual(first[header.index('Part Number')], 'ABC123')

    def test_a_merged_line_says_which_rows_it_absorbed(self):
        from bomlib.report import build_parts_rows
        self.result['rows'][0]['mergedRows'] = [7, 9]
        rows = build_parts_rows(self.result, self.summary)
        self.assertIn('includes rows 7, 9', rows[1][rows[0].index('Notes')])

    def test_the_widths_still_line_up_with_the_columns(self):
        from bomlib.report import build_workbook_sheets
        self.result['rows'][0]['alsoIn'] = [{'name': 'Board B', 'quantity': 250}]
        sheets = build_workbook_sheets(self.result, self.summary)
        parts = next(s for s in sheets if s['name'] == 'Parts')
        self.assertEqual(len(parts['widths']), len(parts['rows'][0]))


class DmsmsCliTests(unittest.TestCase):
    """--dmsms writes the same form the web app builds."""

    def setUp(self):
        self.paths = []
        self._original = bom.build_service
        bom.build_service = self._stub_service

    def tearDown(self):
        bom.build_service = self._original
        for path in self.paths:
            if os.path.exists(path):
                os.unlink(path)

    def _stub_service(self, args):
        return LookupService(clients=[
            StubSupplier('digikey', 'DigiKey', 0.10, lifecycle='Obsolete'),
            StubSupplier('mouser', 'Mouser', 0.20, lifecycle='Active'),
        ], cache=None), None

    def temp(self, suffix):
        handle, path = tempfile.mkstemp(suffix=suffix)
        os.close(handle)
        self.paths.append(path)
        return path

    def source(self):
        path = self.temp('.csv')
        with open(path, 'w', encoding='utf-8') as handle:
            handle.write('Manufacturer Part Number,Qty\nABC123,100\nDEF456,50\n')
        return path

    def run_cli(self, argv):
        errors = io.StringIO()
        original = sys.stderr
        sys.stderr = errors
        try:
            return bom.main(argv), errors.getvalue()
        finally:
            sys.stderr = original

    def test_a_form_is_written_for_the_at_risk_parts(self):
        out = self.temp('.xlsx')
        code, _ = self.run_cli([self.source(), '--dmsms', out, '--program', 'Falcon II',
                                '--quiet', '--no-color'])
        self.assertEqual(code, 0)
        with open(out, 'rb') as handle:
            grid = parse_xlsx(handle.read())
        self.assertEqual(grid[0][0], 'DMSMS Case Form')
        self.assertIn('Falcon II', grid[1][0])
        index = next(i for i, line in enumerate(grid) if line and line[0] == 'Item')
        records = [line for line in grid[index + 1:] if line and str(line[0]).isdigit()]
        self.assertEqual(len(records), 2)

    def test_the_form_needs_a_program_to_be_named_after(self):
        out = self.temp('.xlsx')
        code, errors = self.run_cli([self.source(), '--dmsms', out, '--quiet', '--no-color'])
        self.assertEqual(code, 1)
        self.assertIn('--program', errors)

    def test_naming_a_program_without_a_form_is_refused_rather_than_ignored(self):
        code, errors = self.run_cli([self.source(), '--program', 'Falcon II',
                                     '--quiet', '--no-color'])
        self.assertEqual(code, 1)
        self.assertIn('--dmsms', errors)

    def test_the_status_filter_narrows_what_lands_on_the_form(self):
        out = self.temp('.xlsx')
        code, errors = self.run_cli([self.source(), '--dmsms', out, '--program', 'X',
                                     '--dmsms-status', 'End of Life', '--quiet', '--no-color'])
        # The stubs report Obsolete, so filtering to End of Life leaves nothing.
        self.assertEqual(code, 1)
        self.assertIn('no at-risk parts', errors)
