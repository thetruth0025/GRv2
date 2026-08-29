"""Who can supply each part soonest, and how the report bands them.

The rule under test is a purchasing rule, not a pricing one: stock on hand is
the fastest answer there is, and where several suppliers are equally fast the
cheapest of them wins. Everything else in the report follows from those two.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bomlib import leadtime  # noqa: E402
from bomlib.normalize import Lifecycle  # noqa: E402

SUPPLIERS = [
    {'id': 'digikey', 'name': 'DigiKey'},
    {'id': 'mouser', 'name': 'Mouser'},
    {'id': 'trustedparts', 'name': 'TrustedParts'},
]


def offer(name, days=None, stocked=False, extended=None, unit=None, stock=0, found=True):
    return {
        'supplier': name, 'found': found,
        'stockSufficient': True if stocked else (False if found else None),
        'leadTimeDays': days,
        'leadTimeText': None if days is None else '%d Days' % days,
        'extendedPrice': extended, 'unitPrice': unit, 'stock': stock,
        'supplierPartNumber': name[:2].upper() + '-1', 'orderQuantity': 10,
        'currency': 'USD',
    }


def row(offers, mpn='ABC123', quantity=10, lifecycle=Lifecycle.ACTIVE, **extra):
    entry = {
        'index': 0, 'row': 1, 'mpn': mpn, 'quantity': quantity,
        'manufacturer': 'Acme', 'description': 'A part', 'reference': 'R1',
        'offers': offers,
        'comparison': {'lifecycle': lifecycle, 'lifecycleSeverity': 'ok'},
    }
    entry.update(extra)
    return entry


def summarize(offers, **kwargs):
    return leadtime.summarize_row(row(offers, **kwargs), SUPPLIERS)


class QuickestSupplierTests(unittest.TestCase):
    def test_stock_on_hand_beats_any_quoted_lead_time(self):
        entry = summarize({
            'digikey': offer('DigiKey', days=7, extended=1.0),
            'mouser': offer('Mouser', stocked=True, stock=5000, extended=99.0),
        })
        # Even at a hundred times the price: it is here, and the other is not.
        self.assertEqual(entry['supplier'], 'Mouser')
        self.assertEqual(entry['days'], 0)
        self.assertEqual(entry['availability'], 'In stock')

    def test_the_soonest_quoted_lead_time_wins_when_nobody_has_stock(self):
        entry = summarize({
            'digikey': offer('DigiKey', days=84, extended=1.0),
            'mouser': offer('Mouser', days=28, extended=9.0),
        })
        self.assertEqual(entry['supplier'], 'Mouser')
        self.assertEqual(entry['days'], 28)

    def test_equally_fast_suppliers_are_settled_on_price(self):
        entry = summarize({
            'digikey': offer('DigiKey', days=28, extended=90.0),
            'mouser': offer('Mouser', days=28, extended=40.0),
            'trustedparts': offer('TrustedParts', days=28, extended=65.0),
        })
        self.assertEqual(entry['supplier'], 'Mouser')
        # Named cheapest-first, so the note reads as the ranking it came from.
        self.assertEqual(entry['tiedOn'], ['Mouser', 'TrustedParts', 'DigiKey'])
        self.assertIn('Cheapest of 3', entry['note'])

    def test_several_in_stock_is_also_settled_on_price(self):
        entry = summarize({
            'digikey': offer('DigiKey', stocked=True, stock=100, extended=12.0),
            'mouser': offer('Mouser', stocked=True, stock=100, extended=8.0),
        })
        self.assertEqual(entry['supplier'], 'Mouser')
        self.assertEqual(entry['days'], 0)

    def test_a_supplier_with_no_price_does_not_win_a_tie_over_one_with_a_price(self):
        entry = summarize({
            'digikey': offer('DigiKey', days=28, extended=None),
            'mouser': offer('Mouser', days=28, extended=40.0),
        })
        self.assertEqual(entry['supplier'], 'Mouser')

    def test_a_supplier_that_names_no_date_cannot_win_the_race(self):
        entry = summarize({
            'digikey': offer('DigiKey', days=None, extended=1.0),
            'mouser': offer('Mouser', days=84, extended=99.0),
        })
        self.assertEqual(entry['supplier'], 'Mouser')
        self.assertEqual(entry['days'], 84)

    def test_everyone_who_carries_it_is_listed_behind_the_winner(self):
        entry = summarize({
            'digikey': offer('DigiKey', days=84, extended=1.0),
            'mouser': offer('Mouser', days=28, extended=9.0),
            'trustedparts': offer('TrustedParts', stocked=True, stock=50, extended=50.0),
        })
        self.assertEqual(entry['supplier'], 'TrustedParts')
        self.assertEqual(entry['suppliersCarrying'], 3)
        self.assertEqual([o['supplier'] for o in entry['others']], ['Mouser', 'DigiKey'])
        self.assertEqual([o['availability'] for o in entry['others']],
                         ['4 weeks', '12 weeks'])


class BandTests(unittest.TestCase):
    """The four colours, and the boundaries between them."""

    def band(self, **kwargs):
        return summarize({'digikey': offer('DigiKey', extended=1.0, **kwargs)})['band']

    def test_stock_on_hand_is_the_green_band(self):
        self.assertEqual(self.band(stocked=True, stock=100), leadtime.QUICK)

    def test_inside_three_weeks_is_green_too(self):
        # The gap the colour scheme leaves: a fortnight is not something to
        # plan around, so it sits with stock rather than with the yellow band.
        self.assertEqual(self.band(days=14), leadtime.QUICK)
        self.assertEqual(self.band(days=20), leadtime.QUICK)

    def test_three_to_eight_weeks_is_the_yellow_band(self):
        self.assertEqual(self.band(days=21), leadtime.MEDIUM)
        self.assertEqual(self.band(days=42), leadtime.MEDIUM)
        self.assertEqual(self.band(days=56), leadtime.MEDIUM)

    def test_over_eight_weeks_is_the_orange_band(self):
        self.assertEqual(self.band(days=57), leadtime.LONG)
        self.assertEqual(self.band(days=182), leadtime.LONG)

    def test_nobody_carrying_it_is_the_red_band(self):
        entry = summarize({'digikey': offer('DigiKey', found=False)})
        self.assertEqual(entry['band'], leadtime.NONE)
        self.assertIsNone(entry['supplier'])
        self.assertEqual(entry['suppliersCarrying'], 0)
        # Short in the column, said in full in the note beside it.
        self.assertEqual(entry['availability'], leadtime.NO_SUPPLIER_SHORT)
        self.assertEqual(entry['note'], leadtime.NO_SUPPLIER_TEXT)

    def test_a_row_with_no_offers_at_all_is_also_red(self):
        self.assertEqual(summarize({})['band'], leadtime.NONE)

    def test_carried_with_no_date_is_left_uncoloured_rather_than_guessed_at(self):
        # Claiming a duration nobody quoted would be inventing supply data.
        entry = summarize({'digikey': offer('DigiKey', days=None, extended=1.0)})
        self.assertEqual(entry['band'], leadtime.UNKNOWN)
        self.assertEqual(entry['availability'], leadtime.UNKNOWN_SHORT)
        self.assertIn('none quoting a date', entry['note'])
        self.assertEqual(entry['supplier'], 'DigiKey')


class ReportTests(unittest.TestCase):
    def build(self, *rows):
        return leadtime.build_report({'rows': list(rows), 'suppliers': SUPPLIERS})

    def test_the_worst_lines_come_first(self):
        report = self.build(
            row({'digikey': offer('DigiKey', stocked=True, stock=9, extended=1.0)}, mpn='FINE'),
            row({'digikey': offer('DigiKey', days=90, extended=1.0)}, mpn='SLOW'),
            row({'digikey': offer('DigiKey', found=False)}, mpn='GONE'),
            row({'digikey': offer('DigiKey', days=30, extended=1.0)}, mpn='MID'),
        )
        self.assertEqual([r['mpn'] for r in report['rows']], ['GONE', 'SLOW', 'MID', 'FINE'])

    def test_the_longest_wait_leads_its_own_band(self):
        report = self.build(
            row({'digikey': offer('DigiKey', days=70, extended=1.0)}, mpn='A'),
            row({'digikey': offer('DigiKey', days=200, extended=1.0)}, mpn='B'),
            row({'digikey': offer('DigiKey', days=120, extended=1.0)}, mpn='C'),
        )
        self.assertEqual([r['mpn'] for r in report['rows']], ['B', 'C', 'A'])

    def test_every_band_is_counted(self):
        report = self.build(
            row({'digikey': offer('DigiKey', stocked=True, stock=9, extended=1.0)}, mpn='A'),
            row({'digikey': offer('DigiKey', days=90, extended=1.0)}, mpn='B'),
            row({'digikey': offer('DigiKey', days=30, extended=1.0)}, mpn='C'),
            row({'digikey': offer('DigiKey', found=False)}, mpn='D'),
            row({'digikey': offer('DigiKey', days=None, extended=1.0)}, mpn='E'),
        )
        self.assertEqual(report['counts'], {
            leadtime.NONE: 1, leadtime.LONG: 1, leadtime.MEDIUM: 1,
            leadtime.UNKNOWN: 1, leadtime.QUICK: 1,
        })

    def test_every_looked_up_part_appears_not_just_the_slow_ones(self):
        report = self.build(*[
            row({'digikey': offer('DigiKey', stocked=True, stock=9, extended=1.0)}, mpn='P%d' % i)
            for i in range(5)
        ])
        self.assertEqual(len(report['rows']), 5)


class WorkbookTests(unittest.TestCase):
    """The sheet a buyer opens, and the colours it is read by."""

    def report(self):
        rows = [
            row({'digikey': offer('DigiKey', stocked=True, stock=900, extended=5.0, unit=0.5)},
                mpn='QUICK-1'),
            row({'digikey': offer('DigiKey', days=35, extended=5.0)}, mpn='MID-1'),
            row({'digikey': offer('DigiKey', days=140, extended=5.0)}, mpn='SLOW-1'),
            row({'digikey': offer('DigiKey', found=False)}, mpn='GONE-1'),
        ]
        return leadtime.build_report({'rows': rows, 'suppliers': SUPPLIERS})

    def test_each_band_gets_its_own_fill(self):
        from bomlib.report import build_lead_rows
        from bomlib.xlsx_writer import fill_style
        rows = build_lead_rows(self.report(), styled=True)
        styles = [r[1].style for r in rows[1:]]
        self.assertEqual(styles, [
            fill_style('red'), fill_style('orange'), fill_style('yellow'), fill_style('green'),
        ])

    def test_money_stays_money_inside_a_shaded_row(self):
        from bomlib.report import build_lead_rows, LEAD_COLUMNS
        from bomlib.xlsx_writer import fill_style
        rows = build_lead_rows(self.report(), styled=True)
        column = LEAD_COLUMNS.index('Extended')
        # The last row is the green one, and its price is still formatted.
        self.assertEqual(rows[-1][column].style, fill_style('green', 'money'))
        self.assertEqual(rows[-1][column].value, 5.0)

    def test_the_widths_line_up_with_the_columns(self):
        from bomlib.report import build_lead_workbook_sheets
        for sheet in build_lead_workbook_sheets(self.report(), {'name': 'Board', 'generated': 'now'}):
            self.assertEqual(len(sheet['widths']), len(sheet['rows'][0]), sheet['name'])

    def test_the_legend_counts_what_the_table_holds(self):
        from bomlib.report import build_lead_summary_rows
        flat = [[getattr(c, 'value', c) for c in r]
                for r in build_lead_summary_rows(self.report(), {'name': 'B', 'generated': 'now'})]
        counts = {r[0]: r[1] for r in flat if r[0] in leadtime.BAND_LABEL.values()}
        self.assertEqual(counts['Not available'], 1)
        self.assertEqual(counts['Over 8 weeks'], 1)
        self.assertEqual(counts['3–8 weeks'], 1)
        self.assertEqual(counts['In stock or under 3 weeks'], 1)

    def test_the_file_opens_and_carries_both_sheets(self):
        import io as _io
        from bomlib.report import write_lead_workbook
        from bomlib.spreadsheet import parse_xlsx
        buffer = _io.BytesIO()
        write_lead_workbook(buffer, self.report(), {'name': 'Board', 'generated': 'now'})
        data = buffer.getvalue()
        table = parse_xlsx(data, 'By part')
        self.assertEqual(table[0][1], 'Part Number')
        self.assertEqual([r[1] for r in table[1:]], ['GONE-1', 'SLOW-1', 'MID-1', 'QUICK-1'])
        self.assertIn('Long Lead Times', parse_xlsx(data, 'Lead times')[0][0])

    def test_a_bom_alternate_is_named_on_a_line_nobody_can_supply(self):
        from bomlib.report import build_lead_rows, LEAD_COLUMNS
        entry = row({'digikey': offer('DigiKey', found=False)}, mpn='GONE-1',
                    alternates=[{'mpn': 'SUB-1', 'usable': True}])
        report = leadtime.build_report({'rows': [entry], 'suppliers': SUPPLIERS})
        rows = build_lead_rows(report)
        self.assertIn('BOM alternate available: SUB-1', rows[1][LEAD_COLUMNS.index('Notes')])


if __name__ == '__main__':
    unittest.main()


class NoteTests(unittest.TestCase):
    """The sentence beside each line, which is the only place a reason is given."""

    def note(self, offers, **kwargs):
        return summarize(offers, **kwargs)['note']

    def test_a_line_that_explains_itself_says_nothing(self):
        self.assertEqual(self.note({
            'digikey': offer('DigiKey', days=35, extended=1.0),
            'mouser': offer('Mouser', days=70, extended=1.0),
        }), '')

    def test_a_sole_source_on_a_slow_line_is_called_out(self):
        self.assertEqual(self.note({'digikey': offer('DigiKey', days=120, extended=1.0)}),
                         'Single source — only DigiKey carries it')

    def test_a_sole_source_that_is_quick_is_not_worth_saying(self):
        self.assertEqual(self.note({'digikey': offer('DigiKey', stocked=True, stock=9,
                                                     extended=1.0)}), '')

    def test_an_approved_alternate_rescues_a_line_nobody_carries(self):
        note = self.note({'digikey': offer('DigiKey', found=False)},
                         alternates=[{'mpn': 'SUB-1', 'usable': True}])
        self.assertIn(leadtime.NO_SUPPLIER_TEXT, note)
        self.assertIn('BOM alternate available: SUB-1', note)

    def test_and_a_line_that_is_merely_slow(self):
        note = self.note({'digikey': offer('DigiKey', days=120, extended=1.0)},
                         alternates=[{'mpn': 'SUB-1', 'usable': True}])
        self.assertIn('BOM alternate available: SUB-1', note)

    def test_an_alternate_that_cannot_be_bought_is_not_offered_as_a_rescue(self):
        note = self.note({'digikey': offer('DigiKey', found=False)},
                         alternates=[{'mpn': 'SUB-1', 'usable': False}])
        self.assertEqual(note, leadtime.NO_SUPPLIER_TEXT)

    def test_a_healthy_line_is_not_cluttered_with_its_alternates(self):
        note = self.note({'digikey': offer('DigiKey', stocked=True, stock=9, extended=1.0)},
                         alternates=[{'mpn': 'SUB-1', 'usable': True}])
        self.assertEqual(note, '')


class SummaryExportShadingTests(unittest.TestCase):
    """The availability colours have to reach the summary export too.

    A buyer opening the workbook should not have to cross-reference the
    lead-time sheet to see which lines are the problem, and the two must use
    one palette so they can never disagree.
    """

    def line(self, index, mpn, one):
        """One row, with the comparison the app would really have built."""
        from bomlib.normalize import compare_offers
        entry = row({'digikey': one}, mpn=mpn, quantity=10)
        entry['index'] = index
        entry['comparison'] = compare_offers([one], 10)
        return entry

    def result(self):
        rows = [
            self.line(0, 'QUICK-1', offer('DigiKey', stocked=True, stock=900,
                                          extended=5.0, unit=0.00525)),
            self.line(1, 'MID-1', offer('DigiKey', days=35, stock=900,
                                        extended=5.0, unit=0.5)),
            self.line(2, 'SLOW-1', offer('DigiKey', days=140, stock=900,
                                         extended=5.0, unit=0.5)),
            self.line(3, 'GONE-1', offer('DigiKey', found=False)),
            self.line(4, 'UNKNOWN-1', offer('DigiKey', days=None, stock=900,
                                            extended=5.0, unit=0.5)),
        ]
        return {'rows': rows, 'suppliers': SUPPLIERS, 'stats': {}}

    def summary(self):
        return {'currency': 'USD', 'lines': 5, 'totalQuantity': 50, 'supplierTotals': {},
                'bestMixTotal': 0, 'bestMixLines': 0, 'notFoundLines': 1, 'unpricedLines': 0,
                'errorLines': 0, 'riskLines': [], 'lifecycleCounts': {}}

    def test_the_parts_sheet_is_shaded_by_availability(self):
        from bomlib.report import build_parts_rows, parts_filter_rows
        from bomlib.xlsx_writer import fill_style
        rows = build_parts_rows(self.result(), self.summary(), styled=True)
        data = rows[1:parts_filter_rows(self.result())]
        self.assertEqual([r[1].style for r in data], [
            fill_style('green'), fill_style('yellow'), fill_style('orange'),
            fill_style('red'), fill_style(None),
        ])

    def test_money_keeps_its_format_on_a_shaded_row(self):
        from bomlib.report import build_parts_rows, PARTS_COLUMNS
        from bomlib.xlsx_writer import fill_style
        rows = build_parts_rows(self.result(), self.summary(), styled=True)
        unit = PARTS_COLUMNS.index('Unit Price')
        extended = PARTS_COLUMNS.index('Extended')
        # A sub-cent unit price still needs five decimals, shaded or not.
        self.assertEqual(rows[1][unit].style, fill_style('green', 'money_fine'))
        self.assertEqual(rows[1][unit].value, 0.00525)
        self.assertEqual(rows[1][extended].style, fill_style('green', 'money'))

    def test_quantities_stay_integers_on_a_shaded_row(self):
        from bomlib.report import build_parts_rows, PARTS_COLUMNS
        from bomlib.xlsx_writer import fill_style
        rows = build_parts_rows(self.result(), self.summary(), styled=True)
        self.assertEqual(rows[1][PARTS_COLUMNS.index('Qty')].style, fill_style('green', 'int'))

    def test_an_unstyled_build_is_plain_values_as_before(self):
        from bomlib.report import build_parts_rows
        rows = build_parts_rows(self.result(), self.summary(), styled=False)
        self.assertTrue(all(not hasattr(c, 'style') for row in rows for c in row))

    def test_the_sheet_carries_a_key_to_the_colours(self):
        from bomlib.report import build_parts_rows
        rows = build_parts_rows(self.result(), self.summary(), styled=True)
        flat = [[getattr(c, 'value', c) for c in r] for r in rows]
        self.assertTrue(any(r[0] == 'Colour key' for r in flat))
        labels = [r[0] for r in flat]
        for band in leadtime.BAND_ORDER:
            self.assertIn(leadtime.BAND_LABEL[band], labels)

    def test_the_key_is_left_out_of_the_filter_range(self):
        # Otherwise sorting the table scatters the legend through the parts.
        from bomlib.report import build_parts_rows, parts_filter_rows
        result = self.result()
        rows = build_parts_rows(result, self.summary(), styled=True)
        self.assertGreater(len(rows), parts_filter_rows(result))
        self.assertEqual(parts_filter_rows(result), 1 + len(result['rows']))

    def test_the_report_sheet_shades_the_lines_needing_a_decision(self):
        from bomlib.report import build_report_rows
        rows = build_report_rows(self.result(), self.summary(), {'name': 'B', 'generated': 'now'})
        flat = [[getattr(c, 'value', c) for c in r] for r in rows]
        start = next(i for i, r in enumerate(flat) if str(r[0]).startswith('Needs a decision'))
        from bomlib.xlsx_writer import FILL_STYLES
        shades = [rows[i][0].style for i in range(start + 2, len(rows))
                  if flat[i][0] not in ('', None, 'Colour key')
                  and flat[i][0] not in leadtime.BAND_LABEL.values()]
        colours = {colour for colour, styles in FILL_STYLES.items()
                   for style in shades if style in styles}
        self.assertIn('orange', colours)
        self.assertIn('red', colours)
        self.assertNotIn('green', colours)  # a healthy line is not a decision

    def test_the_summary_and_the_lead_report_share_one_palette(self):
        from bomlib.report import LEAD_FILL, build_lead_rows, build_parts_rows
        from bomlib.xlsx_writer import fill_style
        result = self.result()
        lead = build_lead_rows(leadtime.build_report(result), styled=True)
        parts = build_parts_rows(result, self.summary(), styled=True)
        # The lead report sorts worst first and the parts sheet keeps BOM order,
        # so compare the sets of fills rather than the sequences.
        lead_fills = {r[1].style for r in lead[1:]}
        part_fills = {r[1].style for r in parts[1:1 + len(result['rows'])]}
        self.assertEqual(lead_fills, part_fills)
        self.assertEqual(sorted(LEAD_FILL.values(), key=str),
                         sorted([None, 'green', 'orange', 'red', 'yellow'], key=str))
        self.assertEqual(fill_style(LEAD_FILL[leadtime.QUICK]), fill_style('green'))
