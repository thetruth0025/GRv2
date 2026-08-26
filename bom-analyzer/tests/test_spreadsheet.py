import io
import unittest
import zipfile

from bomlib.spreadsheet import (
    clean_cell,
    detect_delimiter,
    extract_bom,
    find_header_row,
    map_columns,
    parse_delimited,
    parse_workbook,
    parse_xlsx,
)


class DelimitedTests(unittest.TestCase):
    def test_quoted_fields_keep_commas_and_embedded_quotes(self):
        rows = parse_delimited('a,"b,c",d\n1,"say ""hi""",3')
        self.assertEqual(rows[0], ['a', 'b,c', 'd'])
        self.assertEqual(rows[1], ['1', 'say "hi"', '3'])

    def test_quoted_field_may_span_lines(self):
        rows = parse_delimited('ref,desc\nR1,"two\nlines"')
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[1][1], 'two\nlines')

    def test_delimiter_detection_distinguishes_tabs_semicolons_commas(self):
        self.assertEqual(detect_delimiter('a\tb\tc\n1\t2\t3'), '\t')
        self.assertEqual(detect_delimiter('a;b;c\n1;2;3'), ';')
        self.assertEqual(detect_delimiter('a,b,c\n1,2,3'), ',')

    def test_leading_byte_order_mark_does_not_corrupt_first_header(self):
        rows = parse_delimited('﻿Qty,MPN\n5,ABC123')
        self.assertEqual(rows[0][0], 'Qty')

    def test_crlf_parses_the_same_as_lf(self):
        self.assertEqual(parse_delimited('a,b\r\n1,2\r\n'), [['a', 'b'], ['1', '2']])


class MappingTests(unittest.TestCase):
    def test_header_aliases_from_different_eda_tools_map_to_same_fields(self):
        variants = [
            ['Qty', 'Manufacturer Part Number', 'RefDes', 'Manufacturer'],
            ['Quantity', 'MPN', 'Reference', 'Mfr'],
            ['QTY', 'Mfr. Part #', 'Designator', 'Brand'],
            ['Qty Per Board', 'Mfg Part No', 'References', 'Mfg Name'],
        ]
        for headers in variants:
            mapping = map_columns(headers)
            self.assertEqual(mapping.get('quantity'), 0, headers)
            self.assertEqual(mapping.get('mpn'), 1, headers)
            self.assertEqual(mapping.get('reference'), 2, headers)
            self.assertEqual(mapping.get('manufacturer'), 3, headers)

    def test_title_block_above_the_real_header_is_skipped(self):
        rows = [
            ['Acme Widget rev C'],
            ['Generated 2026-01-01'],
            [],
            ['Item', 'Reference', 'Qty', 'Manufacturer Part Number', 'Description'],
            ['1', 'R1', '10', 'RC0603FR-0710KL', 'RES 10K'],
        ]
        self.assertEqual(find_header_row(rows), 3)


class ExtractTests(unittest.TestCase):
    def test_bom_extracts_to_lines_with_quantities_and_references(self):
        csv = '\n'.join([
            'Item,Reference,Qty,Manufacturer,Manufacturer Part Number,Description',
            '1,"C1,C2",300,Murata,GRM188R71H104KA93D,CAP CER 0.1UF',
            '2,R1,500,Yageo,RC0603FR-0710KL,RES 10K',
        ])
        result = extract_bom(parse_delimited(csv))
        self.assertEqual(len(result['lines']), 2)
        self.assertEqual(result['lines'][0]['mpn'], 'GRM188R71H104KA93D')
        self.assertEqual(result['lines'][0]['quantity'], 300)
        self.assertEqual(result['lines'][0]['reference'], 'C1,C2')
        self.assertEqual(result['lines'][0]['manufacturer'], 'Murata')
        self.assertEqual(result['lines'][1]['mpn'], 'RC0603FR-0710KL')

    def test_rows_without_a_part_number_are_skipped_and_counted(self):
        csv = 'Qty,Manufacturer Part Number\n10,ABC123\n5,\n3,DEF456'
        result = extract_bom(parse_delimited(csv))
        self.assertEqual(len(result['lines']), 2)
        self.assertEqual(result['skipped'], 1)

    def test_missing_or_unparseable_quantity_falls_back_to_one(self):
        csv = 'MPN,Qty\nABC123,\nDEF456,n/a\nGHI789,12 pcs'
        result = extract_bom(parse_delimited(csv))
        self.assertEqual(result['lines'][0]['quantity'], 1)
        self.assertEqual(result['lines'][1]['quantity'], 1)
        self.assertEqual(result['lines'][2]['quantity'], 12)

    def test_no_recognizable_headers_yields_nothing_rather_than_garbage(self):
        self.assertEqual(extract_bom(parse_delimited('foo,bar\n1,2'))['lines'], [])


def build_xlsx(headers, rows):
    """Build a minimal but real .xlsx so the reader is exercised against the
    actual OOXML container rather than a stub."""
    shared = []

    def index_of(value):
        if value not in shared:
            shared.append(value)
        return shared.index(value)

    sheet_rows = []
    for row_index, row in enumerate([headers] + rows):
        cells = []
        for col_index, value in enumerate(row):
            ref = chr(65 + col_index) + str(row_index + 1)
            if value == '' or value is None:
                continue
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                cells.append('<c r="%s"><v>%s</v></c>' % (ref, value))
            else:
                cells.append('<c r="%s" t="s"><v>%d</v></c>' % (ref, index_of(str(value))))
        sheet_rows.append('<row r="%d">%s</row>' % (row_index + 1, ''.join(cells)))

    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rel_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pkg_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<worksheet xmlns="%s"><sheetData>%s</sheetData></worksheet>' % (ns, ''.join(sheet_rows))
    )
    shared_xml = (
        '<?xml version="1.0" encoding="UTF-8"?><sst xmlns="%s" count="%d">%s</sst>'
        % (ns, len(shared), ''.join(
            '<si><t>%s</t></si>' % s.replace('&', '&amp;').replace('<', '&lt;') for s in shared))
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="%s" xmlns:r="%s"><sheets>'
        '<sheet name="BOM" sheetId="1" r:id="rId1"/></sheets></workbook>' % (ns, rel_ns)
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="%s">'
        '<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/>'
        '</Relationships>' % pkg_ns
    )

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('xl/workbook.xml', workbook_xml)
        archive.writestr('xl/_rels/workbook.xml.rels', rels_xml)
        archive.writestr('xl/sharedStrings.xml', shared_xml)
        archive.writestr('xl/worksheets/sheet1.xml', sheet_xml)
    return buffer.getvalue()


class XlsxTests(unittest.TestCase):
    def test_reads_back_with_strings_and_numbers_intact(self):
        data = build_xlsx(
            ['Item', 'Reference', 'Qty', 'Manufacturer Part Number', 'Description'],
            [
                [1, 'C1', 300, 'GRM188R71H104KA93D', 'CAP CER 0.1UF'],
                [2, 'R1', 500, 'RC0603FR-0710KL', 'RES 10K'],
            ],
        )
        grid = parse_xlsx(data)
        self.assertEqual(len(grid), 3)
        self.assertEqual(grid[0][3], 'Manufacturer Part Number')
        self.assertEqual(grid[1][2], '300')

        result = extract_bom(grid)
        self.assertEqual(len(result['lines']), 2)
        self.assertEqual(result['lines'][0]['mpn'], 'GRM188R71H104KA93D')
        self.assertEqual(result['lines'][0]['quantity'], 300)
        self.assertEqual(result['lines'][1]['reference'], 'R1')

    def test_recognized_by_contents_even_without_the_extension(self):
        grid = parse_workbook(build_xlsx(['Qty', 'MPN'], [[10, 'ABC123']]), 'upload.bin')
        self.assertEqual(grid[1][1], 'ABC123')

    def test_gaps_in_a_row_keep_remaining_columns_aligned(self):
        data = build_xlsx(
            ['Item', 'Reference', 'Qty', 'Manufacturer Part Number'],
            [[1, '', 25, 'STM32F103C8T6']],
        )
        result = extract_bom(parse_xlsx(data))
        self.assertEqual(result['lines'][0]['mpn'], 'STM32F103C8T6')
        self.assertEqual(result['lines'][0]['quantity'], 25)

    def test_csv_routed_through_parse_workbook_behaves_like_a_csv(self):
        grid = parse_workbook(b'Qty,MPN\n5,ABC123', 'bom.csv')
        self.assertEqual(grid[1], ['5', 'ABC123'])


if __name__ == '__main__':
    unittest.main()


class CleanCellTests(unittest.TestCase):
    """Whitespace and invisible characters around a part number.

    A part number arrives padded far more often than not — copied from a
    datasheet PDF, an ERP export, an indented BOM. Some of that padding is
    ordinary space, and some of it is characters Unicode does not call
    whitespace at all, which are invisible in a spreadsheet and so cannot be
    found and deleted by hand. They are written as escapes here for the same
    reason: in a source file they would be just as invisible.
    """

    # Space characters. str.strip() already knows about all of these.
    SPACES = (' ', '\t', '\xa0', ' ', ' ', '　', ' ')

    # Not whitespace to Unicode, so nothing strips them: zero-width space,
    # byte-order mark, word joiner, soft hyphen, left-to-right mark, and a
    # bidi embedding control.
    HIDDEN = ('​', '﻿', '⁠', '\xad', '‎', '‪')

    def test_ordinary_space_before_and_after_goes(self):
        self.assertEqual(clean_cell('   RC0603FR-0710KL'), 'RC0603FR-0710KL')
        self.assertEqual(clean_cell('RC0603FR-0710KL   '), 'RC0603FR-0710KL')
        self.assertEqual(clean_cell('\tRC0603FR-0710KL\t'), 'RC0603FR-0710KL')

    def test_every_unicode_space_counts_as_space(self):
        for space in self.SPACES:
            self.assertEqual(clean_cell(space + 'ABC123' + space), 'ABC123', repr(space))

    def test_invisible_characters_go_too(self):
        for hidden in self.HIDDEN:
            self.assertEqual(clean_cell(hidden + 'ABC123'), 'ABC123', repr(hidden))
            self.assertEqual(clean_cell('ABC123' + hidden), 'ABC123', repr(hidden))

    def test_an_invisible_character_in_the_middle_goes_as_well(self):
        self.assertEqual(clean_cell('ABC​123'), 'ABC123')

    def test_a_run_of_spaces_inside_collapses_to_one(self):
        self.assertEqual(clean_cell('RES   SMD   10K'), 'RES SMD 10K')

    def test_a_single_interior_space_is_left_alone(self):
        # A few real part numbers carry one, and guessing which would break
        # more than it fixed.
        self.assertEqual(clean_cell('RES SMD 10K'), 'RES SMD 10K')

    def test_a_newline_inside_a_quoted_field_survives(self):
        self.assertEqual(clean_cell('two\nlines'), 'two\nlines')

    def test_a_cell_of_nothing_but_padding_comes_back_empty(self):
        for blank in ('', '   ', '​', '\xa0​ '):
            self.assertEqual(clean_cell(blank), '', repr(blank))

    def test_non_text_values_are_handled(self):
        self.assertEqual(clean_cell(None), '')
        self.assertEqual(clean_cell(42), '42')


class PaddedPartNumberTests(unittest.TestCase):
    """The same padding, arriving through each reader."""

    def csv_mpn(self, raw):
        text = 'Manufacturer Part Number,Qty\n"%s",100\n' % raw
        lines = extract_bom(parse_delimited(text))['lines']
        return lines[0]['mpn'] if lines else None

    def test_a_padded_part_number_in_a_csv_is_read_clean(self):
        for raw in ('   RC0603FR-0710KL', '​RC0603FR-0710KL',
                    '\xa0 RC0603FR-0710KL ', '﻿RC0603FR-0710KL'):
            self.assertEqual(self.csv_mpn(raw), 'RC0603FR-0710KL', repr(raw))

    def test_a_padded_part_number_in_a_workbook_is_read_clean(self):
        from bomlib.xlsx_writer import write_xlsx
        buffer = io.BytesIO()
        write_xlsx(buffer, [
            ['Manufacturer Part Number', 'Qty'],
            ['   RC0603FR-0710KL', 100],
            ['​STM32F103C8T6', 25],
            ['\xa0LM358DR ', 50],
        ])
        lines = extract_bom(parse_workbook(buffer.getvalue(), 'bom.xlsx'))['lines']
        self.assertEqual([line['mpn'] for line in lines],
                         ['RC0603FR-0710KL', 'STM32F103C8T6', 'LM358DR'])

    def test_a_row_padded_into_looking_full_is_still_skipped(self):
        # A cell holding only invisible characters is an empty cell.
        text = 'Manufacturer Part Number,Qty\n"​   ",100\nABC123,5\n'
        parsed = extract_bom(parse_delimited(text))
        self.assertEqual([line['mpn'] for line in parsed['lines']], ['ABC123'])

    def test_padding_does_not_split_one_part_into_two(self):
        # The point of all of it: a padded part and a clean one are the same
        # part, so they merge rather than costing two lookups.
        from bomlib.prepare import prepare_lines
        text = ('Manufacturer Part Number,Qty\n'
                '"  ABC123",10\n'
                '"​ABC123",20\n'
                'ABC123,30\n')
        lines = extract_bom(parse_delimited(text))['lines']
        screened = prepare_lines(lines)
        self.assertEqual([line['mpn'] for line in screened['lines']], ['ABC123'])
        self.assertEqual(screened['lines'][0]['quantity'], 60)

    def test_other_columns_are_cleaned_the_same_way(self):
        text = ('Manufacturer Part Number,Qty,Manufacturer,Description\n'
                'ABC123,5,"​ Yageo ","  RES   SMD  "\n')
        line = extract_bom(parse_delimited(text))['lines'][0]
        self.assertEqual(line['manufacturer'], 'Yageo')
        self.assertEqual(line['description'], 'RES SMD')
