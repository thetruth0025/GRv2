'use strict';

const test = require('node:test');
const assert = require('node:assert');
const zlib = require('zlib');

const {
  parseDelimited,
  detectDelimiter,
  extractBom,
  findHeaderRow,
  mapColumns,
  parseXlsx,
  parseWorkbook,
} = require('../lib/spreadsheet');

test('quoted fields keep their commas and embedded quotes', () => {
  const rows = parseDelimited('a,"b,c",d\n1,"say ""hi""",3');
  assert.deepEqual(rows[0], ['a', 'b,c', 'd']);
  assert.deepEqual(rows[1], ['1', 'say "hi"', '3']);
});

test('a quoted field may span lines', () => {
  const rows = parseDelimited('ref,desc\nR1,"two\nlines"');
  assert.equal(rows.length, 2);
  assert.equal(rows[1][1], 'two\nlines');
});

test('delimiter detection distinguishes tabs, semicolons and commas', () => {
  assert.equal(detectDelimiter('a\tb\tc\n1\t2\t3'), '\t');
  assert.equal(detectDelimiter('a;b;c\n1;2;3'), ';');
  assert.equal(detectDelimiter('a,b,c\n1,2,3'), ',');
});

test('a leading byte-order mark does not corrupt the first header', () => {
  const rows = parseDelimited('﻿Qty,MPN\n5,ABC123');
  assert.equal(rows[0][0], 'Qty');
});

test('CRLF line endings parse the same as LF', () => {
  const rows = parseDelimited('a,b\r\n1,2\r\n');
  assert.deepEqual(rows, [['a', 'b'], ['1', '2']]);
});

test('header aliases from different EDA tools all map to the same fields', () => {
  const variants = [
    ['Qty', 'Manufacturer Part Number', 'RefDes', 'Manufacturer'],
    ['Quantity', 'MPN', 'Reference', 'Mfr'],
    ['QTY', 'Mfr. Part #', 'Designator', 'Brand'],
    ['Qty Per Board', 'Mfg Part No', 'References', 'Mfg Name'],
  ];
  for (const headers of variants) {
    const mapping = mapColumns(headers);
    assert.equal(mapping.quantity, 0, headers.join('|'));
    assert.equal(mapping.mpn, 1, headers.join('|'));
    assert.equal(mapping.reference, 2, headers.join('|'));
    assert.equal(mapping.manufacturer, 3, headers.join('|'));
  }
});

test('a title block above the real header does not become the header', () => {
  const rows = [
    ['Acme Widget rev C'],
    ['Generated 2026-01-01'],
    [],
    ['Item', 'Reference', 'Qty', 'Manufacturer Part Number', 'Description'],
    ['1', 'R1', '10', 'RC0603FR-0710KL', 'RES 10K'],
  ];
  assert.equal(findHeaderRow(rows), 3);
});

test('a BOM extracts to lines with quantities and references', () => {
  const csv = [
    'Item,Reference,Qty,Manufacturer,Manufacturer Part Number,Description',
    '1,"C1,C2",300,Murata,GRM188R71H104KA93D,CAP CER 0.1UF',
    '2,R1,500,Yageo,RC0603FR-0710KL,RES 10K',
  ].join('\n');

  const result = extractBom(parseDelimited(csv));
  assert.equal(result.lines.length, 2);
  assert.equal(result.lines[0].mpn, 'GRM188R71H104KA93D');
  assert.equal(result.lines[0].quantity, 300);
  assert.equal(result.lines[0].reference, 'C1,C2');
  assert.equal(result.lines[0].manufacturer, 'Murata');
  assert.equal(result.lines[1].mpn, 'RC0603FR-0710KL');
});

test('rows without a part number are skipped and counted, not silently dropped', () => {
  const csv = [
    'Qty,Manufacturer Part Number',
    '10,ABC123',
    '5,',
    '3,DEF456',
  ].join('\n');

  const result = extractBom(parseDelimited(csv));
  assert.equal(result.lines.length, 2);
  assert.equal(result.skipped, 1);
});

test('a missing or unparseable quantity falls back to one', () => {
  const csv = ['MPN,Qty', 'ABC123,', 'DEF456,n/a', 'GHI789,12 pcs'].join('\n');
  const result = extractBom(parseDelimited(csv));
  assert.equal(result.lines[0].quantity, 1);
  assert.equal(result.lines[1].quantity, 1);
  assert.equal(result.lines[2].quantity, 12);
});

test('a BOM with no recognizable headers still yields nothing rather than garbage', () => {
  const result = extractBom(parseDelimited('foo,bar\n1,2'));
  assert.equal(result.lines.length, 0);
});

// ── XLSX ────────────────────────────────────────────────────────────────────

// Builds a minimal but real .xlsx (a ZIP of the parts Excel writes) so the
// reader is exercised against the actual container format.
function buildXlsx(headers, rows) {
  const shared = [];
  const indexOf = (value) => {
    const i = shared.indexOf(value);
    if (i !== -1) return i;
    shared.push(value);
    return shared.length - 1;
  };

  const allRows = [headers].concat(rows);
  const sheetRows = allRows
    .map((row, rowIndex) => {
      const cells = row
        .map((value, colIndex) => {
          const ref = String.fromCharCode(65 + colIndex) + (rowIndex + 1);
          if (value === '' || value === null || value === undefined) return '';
          if (typeof value === 'number') {
            return '<c r="' + ref + '"><v>' + value + '</v></c>';
          }
          return '<c r="' + ref + '" t="s"><v>' + indexOf(String(value)) + '</v></c>';
        })
        .join('');
      return '<row r="' + (rowIndex + 1) + '">' + cells + '</row>';
    })
    .join('');

  const sheetXml =
    '<?xml version="1.0" encoding="UTF-8"?><worksheet><sheetData>' + sheetRows + '</sheetData></worksheet>';
  const sharedXml =
    '<?xml version="1.0" encoding="UTF-8"?><sst count="' + shared.length + '">' +
    shared.map((s) => '<si><t>' + s.replace(/&/g, '&amp;').replace(/</g, '&lt;') + '</t></si>').join('') +
    '</sst>';
  const workbookXml =
    '<?xml version="1.0" encoding="UTF-8"?><workbook><sheets>' +
    '<sheet name="BOM" sheetId="1" r:id="rId1"/></sheets></workbook>';
  const relsXml =
    '<?xml version="1.0" encoding="UTF-8"?><Relationships>' +
    '<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/></Relationships>';

  return zipFiles([
    ['xl/workbook.xml', workbookXml],
    ['xl/_rels/workbook.xml.rels', relsXml],
    ['xl/sharedStrings.xml', sharedXml],
    ['xl/worksheets/sheet1.xml', sheetXml],
  ]);
}

function crc32(buffer) {
  let crc = ~0;
  for (let i = 0; i < buffer.length; i++) {
    crc ^= buffer[i];
    for (let bit = 0; bit < 8; bit++) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return ~crc >>> 0;
}

function zipFiles(files) {
  const locals = [];
  const centrals = [];
  let offset = 0;

  for (const [name, content] of files) {
    const nameBuf = Buffer.from(name, 'utf8');
    const raw = Buffer.from(content, 'utf8');
    const deflated = zlib.deflateRawSync(raw);
    const checksum = crc32(raw);

    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt32LE(0, 10);
    local.writeUInt32LE(checksum, 14);
    local.writeUInt32LE(deflated.length, 18);
    local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    locals.push(local, nameBuf, deflated);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0, 8);
    central.writeUInt16LE(8, 10);
    central.writeUInt32LE(0, 12);
    central.writeUInt32LE(checksum, 16);
    central.writeUInt32LE(deflated.length, 20);
    central.writeUInt32LE(raw.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, nameBuf);

    offset += local.length + nameBuf.length + deflated.length;
  }

  const centralBuf = Buffer.concat(centrals);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(files.length, 8);
  eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(centralBuf.length, 12);
  eocd.writeUInt32LE(offset, 16);

  return Buffer.concat([Buffer.concat(locals), centralBuf, eocd]);
}

test('an .xlsx BOM reads back with its strings and numbers intact', () => {
  const buffer = buildXlsx(
    ['Item', 'Reference', 'Qty', 'Manufacturer Part Number', 'Description'],
    [
      [1, 'C1', 300, 'GRM188R71H104KA93D', 'CAP CER 0.1UF'],
      [2, 'R1', 500, 'RC0603FR-0710KL', 'RES 10K'],
    ]
  );

  const grid = parseXlsx(buffer);
  assert.equal(grid.length, 3);
  assert.equal(grid[0][3], 'Manufacturer Part Number');
  assert.equal(grid[1][2], '300');

  const result = extractBom(grid);
  assert.equal(result.lines.length, 2);
  assert.equal(result.lines[0].mpn, 'GRM188R71H104KA93D');
  assert.equal(result.lines[0].quantity, 300);
  assert.equal(result.lines[1].reference, 'R1');
});

test('an .xlsx is recognized by its contents even without the extension', () => {
  const buffer = buildXlsx(['Qty', 'MPN'], [[10, 'ABC123']]);
  const grid = parseWorkbook(buffer, 'upload.bin');
  assert.equal(grid[1][1], 'ABC123');
});

test('gaps in a spreadsheet row keep the remaining columns aligned', () => {
  const buffer = buildXlsx(
    ['Item', 'Reference', 'Qty', 'Manufacturer Part Number'],
    [[1, '', 25, 'STM32F103C8T6']]
  );
  const result = extractBom(parseXlsx(buffer));
  assert.equal(result.lines[0].mpn, 'STM32F103C8T6');
  assert.equal(result.lines[0].quantity, 25);
});

test('a CSV routed through parseWorkbook behaves like a CSV', () => {
  const grid = parseWorkbook(Buffer.from('Qty,MPN\n5,ABC123'), 'bom.csv');
  assert.deepEqual(grid[1], ['5', 'ABC123']);
});
