"""Write a styled single-sheet .xlsx using zipfile and hand-built XML.

The counterpart to the reader in spreadsheet.py, and dependency-free for the
same reason. Deliberately minimal: one sheet, inline strings, a small fixed
style table — enough for a purchasing deliverable that opens cleanly in Excel,
LibreOffice and Numbers, and nothing more.
"""

import re
import zipfile

# Style indices into the cellXfs table built by _styles_xml().
STYLE_DEFAULT = 0
STYLE_HEADER = 1
STYLE_MONEY = 2
STYLE_MONEY_FINE = 3
STYLE_INT = 4
STYLE_BAD = 5
STYLE_WARN = 6
STYLE_GOOD = 7
STYLE_MUTED = 8

# Excel rejects most control characters outright.
_ILLEGAL = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')


def column_letter(index):
    """0 → A, 25 → Z, 26 → AA."""
    letters = ''
    index += 1
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _escape(text):
    return (
        _ILLEGAL.sub('', str(text))
        .replace('&', '&amp;')
        .replace('<', '&lt;')
        .replace('>', '&gt;')
    )


class Cell:
    """A value plus the style to render it with."""

    __slots__ = ('value', 'style')

    def __init__(self, value, style=STYLE_DEFAULT):
        self.value = value
        self.style = style


def _cell_xml(ref, cell):
    value = cell.value
    style = cell.style

    if value is None or value == '':
        # An empty styled cell still has to exist so banding and borders line up.
        return '<c r="%s" s="%d"/>' % (ref, style)
    if isinstance(value, bool):
        return '<c r="%s" s="%d" t="b"><v>%d</v></c>' % (ref, style, 1 if value else 0)
    if isinstance(value, (int, float)):
        return '<c r="%s" s="%d"><v>%s</v></c>' % (ref, style, repr(value) if isinstance(value, float) else value)
    return '<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
        ref, style, _escape(value)
    )


def _sheet_xml(rows, widths, freeze_rows, autofilter):
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

    pane = ''
    if freeze_rows:
        pane = (
            '<sheetView workbookViewId="0"><pane ySplit="%d" topLeftCell="A%d" '
            'activePane="bottomLeft" state="frozen"/></sheetView>'
            % (freeze_rows, freeze_rows + 1)
        )
    else:
        pane = '<sheetView workbookViewId="0"/>'

    cols = ''
    if widths:
        entries = ''.join(
            '<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i + 1, i + 1, width)
            for i, width in enumerate(widths)
        )
        cols = '<cols>%s</cols>' % entries

    body = []
    for row_index, row in enumerate(rows):
        cells = ''.join(
            _cell_xml(column_letter(col_index) + str(row_index + 1), cell)
            for col_index, cell in enumerate(row)
        )
        body.append('<row r="%d">%s</row>' % (row_index + 1, cells))

    filter_xml = ''
    if autofilter and rows:
        filter_xml = '<autoFilter ref="A1:%s%d"/>' % (column_letter(len(rows[0]) - 1), len(rows))

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="%s"><sheetViews>%s</sheetViews>%s'
        '<sheetData>%s</sheetData>%s</worksheet>'
        % (ns, pane, cols, ''.join(body), filter_xml)
    )


def _styles_xml():
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="%s">'
        '<numFmts count="2">'
        '<numFmt numFmtId="164" formatCode="#,##0.00"/>'
        '<numFmt numFmtId="165" formatCode="#,##0.00000"/>'
        '</numFmts>'
        '<fonts count="5">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
        '<font><sz val="11"/><color rgb="FFC00000"/><name val="Calibri"/></font>'
        '<font><sz val="11"/><color rgb="FF9C6500"/><name val="Calibri"/></font>'
        '<font><sz val="11"/><color rgb="FF808080"/><name val="Calibri"/></font>'
        '</fonts>'
        '<fills count="3">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3B4D"/>'
        '<bgColor indexed="64"/></patternFill></fill>'
        '</fills>'
        '<borders count="2">'
        '<border><left/><right/><top/><bottom/><diagonal/></border>'
        '<border><left/><right/><top/><bottom style="thin">'
        '<color rgb="FF8EA9BB"/></bottom><diagonal/></border>'
        '</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="9">'
        # 0 default
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        # 1 header
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" '
        'applyFill="1" applyBorder="1" applyAlignment="1">'
        '<alignment vertical="center" wrapText="1"/></xf>'
        # 2 money (2dp)
        '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        # 3 money (5dp, for sub-cent unit prices)
        '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        # 4 integer with thousands separator
        '<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        # 5 bad, 6 warn, 7 good, 8 muted
        '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>' % ns
    )


def _content_types_xml(sheet_count):
    ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
    doc = 'application/vnd.openxmlformats-officedocument.spreadsheetml'
    sheets = ''.join(
        '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="%s.worksheet+xml"/>'
        % (i + 1, doc) for i in range(sheet_count)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="%s">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="%s.sheet.main+xml"/>'
        '%s'
        '<Override PartName="/xl/styles.xml" ContentType="%s.styles+xml"/>'
        '</Types>' % (ns, doc, sheets, doc)
    )


def write_xlsx(path, sheets, sheet_name='BOM', widths=None, freeze_rows=1, autofilter=True):
    """Write one or more sheets to a workbook.

    `sheets` is either a list of rows (a single sheet, named by `sheet_name`)
    or a list of {'name', 'rows', 'widths'} dicts for a multi-sheet workbook.
    """
    if sheets and isinstance(sheets[0], dict) and 'rows' in sheets[0]:
        specs = sheets
    else:
        specs = [{'name': sheet_name, 'rows': sheets, 'widths': widths}]

    prepared = []
    for index, spec in enumerate(specs):
        rows = [
            [c if isinstance(c, Cell) else Cell(c) for c in row]
            for row in (spec.get('rows') or [])
        ]
        width = max((len(row) for row in rows), default=0)
        for row in rows:
            while len(row) < width:
                row.append(Cell(None))
        prepared.append({
            'name': _escape(spec.get('name') or ('Sheet%d' % (index + 1)))[:31] or ('Sheet%d' % (index + 1)),
            'rows': rows,
            'widths': spec.get('widths'),
        })

    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rel_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pkg_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'

    sheet_tags = ''.join(
        '<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (spec['name'], i + 1, i + 1)
        for i, spec in enumerate(prepared)
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="%s" xmlns:r="%s"><sheets>%s</sheets></workbook>'
        % (ns, rel_ns, sheet_tags)
    )

    sheet_rels = ''.join(
        '<Relationship Id="rId%d" Type="%s/worksheet" Target="worksheets/sheet%d.xml"/>'
        % (i + 1, rel_ns, i + 1) for i in range(len(prepared))
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="%s">%s'
        '<Relationship Id="rId%d" Type="%s/styles" Target="styles.xml"/>'
        '</Relationships>' % (pkg_ns, sheet_rels, len(prepared) + 1, rel_ns)
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="%s">'
        '<Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>' % (pkg_ns, rel_ns)
    )

    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('[Content_Types].xml', _content_types_xml(len(prepared)))
        archive.writestr('_rels/.rels', root_rels)
        archive.writestr('xl/workbook.xml', workbook)
        archive.writestr('xl/_rels/workbook.xml.rels', workbook_rels)
        archive.writestr('xl/styles.xml', _styles_xml())
        for index, spec in enumerate(prepared):
            archive.writestr(
                'xl/worksheets/sheet%d.xml' % (index + 1),
                _sheet_xml(spec['rows'], spec['widths'], freeze_rows, autofilter),
            )
    return path
