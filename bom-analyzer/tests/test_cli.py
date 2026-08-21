import argparse
import json
import os
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

    def test_the_workbook_opens_and_keeps_numbers_numeric(self):
        path = self.temp('.xlsx')
        write_workbook(path, self.result, self.summary)
        with open(path, 'rb') as handle:
            grid = parse_xlsx(handle.read())
        self.assertEqual(grid[0][1], 'Part Number')
        self.assertEqual(grid[1][1], 'ABC123')
        self.assertEqual(grid[1][2], '100')

    def test_json_carries_the_full_result_for_scripting(self):
        path = self.temp('.json')
        write_json(path, self.result, self.summary)
        with open(path, encoding='utf-8') as handle:
            data = json.load(handle)
        self.assertEqual(sorted(data), ['rows', 'stats', 'summary', 'suppliers'])
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
            grid = parse_xlsx(handle.read())
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
