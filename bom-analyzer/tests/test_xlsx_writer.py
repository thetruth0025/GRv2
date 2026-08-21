import os
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

from bomlib.spreadsheet import parse_xlsx
from bomlib.xlsx_writer import (
    Cell,
    STYLE_BAD,
    STYLE_HEADER,
    STYLE_INT,
    STYLE_MONEY_FINE,
    column_letter,
    write_xlsx,
)


class ColumnLetterTests(unittest.TestCase):
    def test_counts_past_z_the_way_spreadsheets_do(self):
        self.assertEqual(column_letter(0), 'A')
        self.assertEqual(column_letter(25), 'Z')
        self.assertEqual(column_letter(26), 'AA')
        self.assertEqual(column_letter(27), 'AB')
        self.assertEqual(column_letter(51), 'AZ')
        self.assertEqual(column_letter(52), 'BA')


class WriteTests(unittest.TestCase):
    def setUp(self):
        handle, self.path = tempfile.mkstemp(suffix='.xlsx')
        os.close(handle)

    def tearDown(self):
        os.unlink(self.path)

    def _bytes(self):
        with open(self.path, 'rb') as handle:
            return handle.read()

    def test_values_survive_a_round_trip_through_the_reader(self):
        write_xlsx(self.path, [
            [Cell('Part', STYLE_HEADER), Cell('Qty', STYLE_HEADER), Cell('Price', STYLE_HEADER)],
            [Cell('ABC123'), Cell(100, STYLE_INT), Cell(1.2345, STYLE_MONEY_FINE)],
            [Cell('DEF456'), Cell(5, STYLE_INT), Cell(0.00784, STYLE_MONEY_FINE)],
        ])
        grid = parse_xlsx(self._bytes())
        self.assertEqual(grid[0], ['Part', 'Qty', 'Price'])
        self.assertEqual(grid[1][0], 'ABC123')
        self.assertEqual(grid[1][1], '100')
        self.assertEqual(grid[1][2], '1.2345')
        self.assertEqual(grid[2][2], '0.00784')

    def test_the_package_contains_every_part_a_reader_expects(self):
        write_xlsx(self.path, [[Cell('a')]])
        with zipfile.ZipFile(self.path) as archive:
            names = archive.namelist()
            for required in ('[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml',
                             'xl/_rels/workbook.xml.rels', 'xl/styles.xml',
                             'xl/worksheets/sheet1.xml'):
                self.assertIn(required, names)
            # OPC requires the content-type map to lead the archive.
            self.assertEqual(names[0], '[Content_Types].xml')
            for name in names:
                ET.fromstring(archive.read(name))

    def test_markup_in_a_value_is_escaped_not_injected(self):
        write_xlsx(self.path, [[Cell('R&D <tag> "quoted"')]])
        grid = parse_xlsx(self._bytes())
        self.assertEqual(grid[0][0], 'R&D <tag> "quoted"')

    def test_control_characters_are_stripped_rather_than_corrupting_the_file(self):
        write_xlsx(self.path, [[Cell('bad\x00value\x07here')]])
        with zipfile.ZipFile(self.path) as archive:
            ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
        self.assertEqual(parse_xlsx(self._bytes())[0][0], 'badvaluehere')

    def test_ragged_rows_are_padded_so_every_row_has_the_same_width(self):
        write_xlsx(self.path, [[Cell('a'), Cell('b'), Cell('c')], [Cell('d')]])
        with zipfile.ZipFile(self.path) as archive:
            xml = archive.read('xl/worksheets/sheet1.xml').decode()
        # The short row still declares cells up to column C.
        self.assertIn('r="C2"', xml)

    def test_freeze_panes_widths_and_autofilter_are_emitted(self):
        write_xlsx(self.path, [[Cell('a'), Cell('b')], [Cell(1), Cell(2)]],
                   widths=[20, 8], freeze_rows=1, autofilter=True)
        with zipfile.ZipFile(self.path) as archive:
            xml = archive.read('xl/worksheets/sheet1.xml').decode()
        self.assertIn('state="frozen"', xml)
        self.assertIn('ySplit="1"', xml)
        self.assertIn('width="20"', xml)
        self.assertIn('<autoFilter ref="A1:B2"/>', xml)

    def test_numbers_stay_numeric_and_text_stays_text(self):
        write_xlsx(self.path, [[Cell(42, STYLE_INT), Cell('42'), Cell(0.5)]])
        with zipfile.ZipFile(self.path) as archive:
            xml = archive.read('xl/worksheets/sheet1.xml').decode()
        # A numeric cell has no t= attribute; a string cell is an inline string.
        self.assertIn('<c r="A1" s="4"><v>42</v></c>', xml)
        self.assertIn('t="inlineStr"', xml)

    def test_a_long_sheet_name_is_truncated_to_what_excel_accepts(self):
        write_xlsx(self.path, [[Cell('a')]], sheet_name='x' * 60)
        with zipfile.ZipFile(self.path) as archive:
            workbook = archive.read('xl/workbook.xml').decode()
        name = workbook.split('name="')[1].split('"')[0]
        self.assertEqual(len(name), 31)

    def test_a_styled_cell_keeps_its_style_index(self):
        write_xlsx(self.path, [[Cell('Obsolete', STYLE_BAD)]])
        with zipfile.ZipFile(self.path) as archive:
            xml = archive.read('xl/worksheets/sheet1.xml').decode()
        self.assertIn('s="%d"' % STYLE_BAD, xml)


if __name__ == '__main__':
    unittest.main()
