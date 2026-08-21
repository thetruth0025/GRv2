'use strict';

const zlib = require('zlib');

// Reads the file formats BOMs actually arrive in — CSV/TSV exported from an
// EDA tool, or the .xlsx a purchasing department sends. Implemented directly
// rather than with a spreadsheet library so the project stays dependency-free
// and can be audited end to end.

// ── Delimited text ─────────────────────────────────────────────────────────

function detectDelimiter(text) {
  const sample = text.slice(0, 8192).split(/\r?\n/).slice(0, 20);
  const candidates = [',', '\t', ';', '|'];
  let best = ',';
  let bestScore = -1;
  for (const delimiter of candidates) {
    const counts = sample
      .filter((line) => line.trim())
      .map((line) => countOutsideQuotes(line, delimiter));
    if (counts.length === 0) continue;
    const total = counts.reduce((a, b) => a + b, 0);
    if (total === 0) continue;
    const mean = total / counts.length;
    const variance = counts.reduce((sum, c) => sum + Math.pow(c - mean, 2), 0) / counts.length;
    // Favour the delimiter that appears often AND consistently per line.
    const score = mean - variance;
    if (score > bestScore) {
      bestScore = score;
      best = delimiter;
    }
  }
  return best;
}

function countOutsideQuotes(line, delimiter) {
  let count = 0;
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') i++;
      else inQuotes = !inQuotes;
    } else if (ch === delimiter && !inQuotes) {
      count++;
    }
  }
  return count;
}

function parseDelimited(text, delimiter) {
  const sep = delimiter || detectDelimiter(text);
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;

  const clean = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;

  for (let i = 0; i < clean.length; i++) {
    const ch = clean[i];
    if (inQuotes) {
      if (ch === '"') {
        if (clean[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
      continue;
    }
    if (ch === '"') {
      inQuotes = true;
    } else if (ch === sep) {
      row.push(field);
      field = '';
    } else if (ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else if (ch === '\r') {
      // Consumed with the following \n; a lone \r also ends the record.
      if (clean[i + 1] !== '\n') {
        row.push(field);
        rows.push(row);
        row = [];
        field = '';
      }
    } else {
      field += ch;
    }
  }
  if (field !== '' || row.length > 0) {
    row.push(field);
    rows.push(row);
  }

  return rows.map((r) => r.map((cell) => cell.trim()));
}

// ── XLSX ───────────────────────────────────────────────────────────────────

const EOCD_SIG = 0x06054b50;
const CENTRAL_SIG = 0x02014b50;
const LOCAL_SIG = 0x04034b50;

function readZipEntries(buffer) {
  const eocd = findEocd(buffer);
  if (eocd === -1) throw new Error('Not a valid .xlsx file (no ZIP end-of-directory record)');

  const entryCount = buffer.readUInt16LE(eocd + 10);
  let offset = buffer.readUInt32LE(eocd + 16);
  const entries = new Map();

  for (let i = 0; i < entryCount; i++) {
    if (offset + 46 > buffer.length) break;
    if (buffer.readUInt32LE(offset) !== CENTRAL_SIG) break;
    const method = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const nameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localOffset = buffer.readUInt32LE(offset + 42);
    const name = buffer.toString('utf8', offset + 46, offset + 46 + nameLength);
    entries.set(name, { method, compressedSize, localOffset });
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

function findEocd(buffer) {
  const min = Math.max(0, buffer.length - 65557);
  for (let i = buffer.length - 22; i >= min; i--) {
    if (buffer.readUInt32LE(i) === EOCD_SIG) return i;
  }
  return -1;
}

function readZipFile(buffer, entries, name) {
  const entry = entries.get(name);
  if (!entry) return null;
  const { localOffset } = entry;
  if (buffer.readUInt32LE(localOffset) !== LOCAL_SIG) return null;
  const nameLength = buffer.readUInt16LE(localOffset + 26);
  const extraLength = buffer.readUInt16LE(localOffset + 28);
  const start = localOffset + 30 + nameLength + extraLength;
  const data = buffer.subarray(start, start + entry.compressedSize);
  if (entry.method === 0) return data.toString('utf8');
  if (entry.method === 8) return zlib.inflateRawSync(data).toString('utf8');
  throw new Error('Unsupported ZIP compression method ' + entry.method + ' for ' + name);
}

function decodeXmlEntities(text) {
  return text.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (match, entity) => {
    if (entity[0] === '#') {
      const code = entity[1] === 'x' || entity[1] === 'X'
        ? parseInt(entity.slice(2), 16)
        : parseInt(entity.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : match;
    }
    switch (entity) {
      case 'amp': return '&';
      case 'lt': return '<';
      case 'gt': return '>';
      case 'quot': return '"';
      case 'apos': return "'";
      default: return match;
    }
  });
}

function parseSharedStrings(xml) {
  if (!xml) return [];
  const strings = [];
  const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
  let match;
  while ((match = siRe.exec(xml)) !== null) {
    // A shared string can be split across several runs (<r><t>…</t></r>).
    const parts = [];
    const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
    let tMatch;
    while ((tMatch = tRe.exec(match[1])) !== null) parts.push(decodeXmlEntities(tMatch[1]));
    strings.push(parts.join(''));
  }
  return strings;
}

function columnToIndex(ref) {
  const letters = String(ref || '').match(/^[A-Z]+/i);
  if (!letters) return null;
  let index = 0;
  const upper = letters[0].toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

function parseWorksheet(xml, sharedStrings) {
  const rows = [];
  if (!xml) return rows;
  const rowRe = /<row\b[^>]*>([\s\S]*?)<\/row>|<row\b[^>]*\/>/g;
  let rowMatch;

  while ((rowMatch = rowRe.exec(xml)) !== null) {
    const inner = rowMatch[1] || '';
    const cells = [];
    const cellRe = /<c\b([^>]*)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let cellMatch;

    while ((cellMatch = cellRe.exec(inner)) !== null) {
      const attrs = cellMatch[1] || '';
      const body = cellMatch[2] || '';
      const refMatch = attrs.match(/\br="([A-Z]+\d+)"/i);
      const typeMatch = attrs.match(/\bt="([^"]+)"/i);
      const type = typeMatch ? typeMatch[1] : 'n';
      const index = refMatch ? columnToIndex(refMatch[1]) : cells.length;

      let value = '';
      if (type === 'inlineStr') {
        const parts = [];
        const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
        let tMatch;
        while ((tMatch = tRe.exec(body)) !== null) parts.push(decodeXmlEntities(tMatch[1]));
        value = parts.join('');
      } else {
        const vMatch = body.match(/<v\b[^>]*>([\s\S]*?)<\/v>/);
        const raw = vMatch ? decodeXmlEntities(vMatch[1]) : '';
        if (type === 's') {
          const idx = parseInt(raw, 10);
          value = Number.isFinite(idx) && sharedStrings[idx] !== undefined ? sharedStrings[idx] : '';
        } else if (type === 'b') {
          value = raw === '1' ? 'TRUE' : 'FALSE';
        } else {
          value = raw;
        }
      }

      if (index === null || index < 0) continue;
      while (cells.length < index) cells.push('');
      cells[index] = String(value).trim();
    }
    rows.push(cells);
  }
  return rows;
}

function resolveFirstSheetPath(buffer, entries) {
  const workbook = readZipFile(buffer, entries, 'xl/workbook.xml');
  const rels = readZipFile(buffer, entries, 'xl/_rels/workbook.xml.rels');

  if (workbook && rels) {
    const sheetMatch = workbook.match(/<sheet\b[^>]*\/>/);
    if (sheetMatch) {
      const idMatch = sheetMatch[0].match(/r:id="([^"]+)"/);
      if (idMatch) {
        const relRe = new RegExp('<Relationship\\b[^>]*Id="' + idMatch[1] + '"[^>]*>');
        const relMatch = rels.match(relRe);
        if (relMatch) {
          const targetMatch = relMatch[0].match(/Target="([^"]+)"/);
          if (targetMatch) {
            let target = targetMatch[1].replace(/^\/?xl\//, '').replace(/^\//, '');
            const path = 'xl/' + target;
            if (entries.has(path)) return path;
          }
        }
      }
    }
  }

  for (const name of entries.keys()) {
    if (/^xl\/worksheets\/sheet[^/]*\.xml$/.test(name)) return name;
  }
  return null;
}

function parseXlsx(buffer) {
  const entries = readZipEntries(buffer);
  const sheetPath = resolveFirstSheetPath(buffer, entries);
  if (!sheetPath) throw new Error('No worksheet found inside the .xlsx file');
  const sharedStrings = parseSharedStrings(readZipFile(buffer, entries, 'xl/sharedStrings.xml'));
  return parseWorksheet(readZipFile(buffer, entries, sheetPath), sharedStrings);
}

// ── Header detection and column mapping ────────────────────────────────────

// Each field lists header aliases in descending confidence. Matching is done on
// a squashed form so "Mfr. Part #" and "mfr_part_no" both land on `mpn`.
const FIELD_ALIASES = {
  mpn: [
    'manufacturerpartnumber', 'mfrpartnumber', 'mfgpartnumber', 'manufacturerpartno',
    'mfrpartno', 'mfgpartno', 'manufacturerpart', 'mfrpart', 'mfgpart', 'mpn',
    'manufacturernumber', 'partnumber', 'partno', 'partnum', 'part', 'componentpartnumber',
    'vendorpartnumber', 'orderingcode', 'ordernumber',
  ],
  quantity: ['quantity', 'qty', 'qtyper', 'quantityper', 'qtyperboard', 'quantityperboard', 'count', 'amount', 'qtyrequired'],
  reference: ['referencedesignator', 'referencedesignators', 'reference', 'references', 'refdes', 'refdesignator', 'designator', 'designators', 'ref'],
  manufacturer: ['manufacturer', 'manufacturername', 'mfr', 'mfg', 'mfrname', 'mfgname', 'brand', 'vendor', 'supplier', 'make'],
  description: ['description', 'desc', 'partdescription', 'componentdescription', 'comment', 'value', 'name', 'partname'],
  footprint: ['footprint', 'package', 'packagetype', 'casecode', 'case'],
};

function squash(text) {
  return String(text || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function scoreHeader(header, field) {
  const key = squash(header);
  if (!key) return 0;
  const aliases = FIELD_ALIASES[field] || [];
  for (let i = 0; i < aliases.length; i++) {
    const alias = aliases[i];
    // Later aliases are weaker matches, so their score decays.
    const weight = 100 - i * 2;
    if (key === alias) return weight;
    if (key.startsWith(alias) || key.endsWith(alias)) return Math.round(weight * 0.8);
    if (key.includes(alias)) return Math.round(weight * 0.6);
  }
  return 0;
}

// BOM exports often carry a title block or blank rows above the real header, so
// the header row is found by scoring rows rather than assuming row 0.
function findHeaderRow(rows) {
  const limit = Math.min(rows.length, 25);
  let bestIndex = 0;
  let bestScore = -1;

  for (let i = 0; i < limit; i++) {
    const row = rows[i] || [];
    const filled = row.filter((cell) => String(cell || '').trim()).length;
    if (filled < 2) continue;
    let score = 0;
    for (const field of Object.keys(FIELD_ALIASES)) {
      let fieldBest = 0;
      for (const cell of row) fieldBest = Math.max(fieldBest, scoreHeader(cell, field));
      score += fieldBest;
    }
    // A header row is text, not data; numeric cells argue against it.
    const numeric = row.filter((cell) => /^-?\d+(\.\d+)?$/.test(String(cell || '').trim())).length;
    score -= numeric * 15;
    score += filled;
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }
  return bestScore > 40 ? bestIndex : 0;
}

function mapColumns(headers) {
  const mapping = {};
  const taken = new Set();

  // Resolve the strongest header/field pairs first so a single "Part Number"
  // column is not claimed by a weaker field.
  const pairs = [];
  for (const field of Object.keys(FIELD_ALIASES)) {
    headers.forEach((header, index) => {
      const score = scoreHeader(header, field);
      if (score > 0) pairs.push({ field, index, score });
    });
  }
  pairs.sort((a, b) => b.score - a.score);

  for (const pair of pairs) {
    if (mapping[pair.field] !== undefined) continue;
    if (taken.has(pair.index)) continue;
    mapping[pair.field] = pair.index;
    taken.add(pair.index);
  }
  return mapping;
}

function parseWorkbook(input, filename) {
  const name = String(filename || '').toLowerCase();
  if (Buffer.isBuffer(input) && (name.endsWith('.xlsx') || name.endsWith('.xlsm') || isZip(input))) {
    return parseXlsx(input);
  }
  const text = Buffer.isBuffer(input) ? input.toString('utf8') : String(input);
  const delimiter = name.endsWith('.tsv') ? '\t' : undefined;
  return parseDelimited(text, delimiter);
}

function isZip(buffer) {
  return buffer.length > 4 && buffer.readUInt32LE(0) === LOCAL_SIG;
}

// Turns a raw grid into BOM lines plus the mapping decisions, which the UI
// shows so the user can correct a bad guess instead of re-exporting.
function extractBom(rows) {
  const grid = (rows || []).filter((row) => Array.isArray(row));
  if (grid.length === 0) {
    return { headerRow: 0, headers: [], mapping: {}, lines: [], skipped: 0 };
  }

  const headerRow = findHeaderRow(grid);
  const headers = (grid[headerRow] || []).map((cell) => String(cell || '').trim());
  const mapping = mapColumns(headers);
  const lines = [];
  let skipped = 0;

  for (let i = headerRow + 1; i < grid.length; i++) {
    const row = grid[i];
    if (!row || row.every((cell) => !String(cell || '').trim())) continue;
    const line = lineFromRow(row, mapping, i);
    if (!line.mpn) {
      skipped++;
      continue;
    }
    lines.push(line);
  }

  return { headerRow, headers, mapping, lines, skipped, totalRows: grid.length };
}

function lineFromRow(row, mapping, rowIndex) {
  const cell = (field) => {
    const index = mapping[field];
    if (index === undefined) return '';
    return String(row[index] === undefined ? '' : row[index]).trim();
  };

  const quantityRaw = cell('quantity');
  const quantity = parseInt(String(quantityRaw).replace(/[^0-9-]/g, ''), 10);

  return {
    row: rowIndex + 1,
    mpn: cell('mpn'),
    quantity: Number.isFinite(quantity) && quantity > 0 ? quantity : 1,
    quantityRaw: quantityRaw || null,
    reference: cell('reference') || null,
    manufacturer: cell('manufacturer') || null,
    description: cell('description') || null,
    footprint: cell('footprint') || null,
  };
}

module.exports = {
  parseDelimited,
  detectDelimiter,
  parseXlsx,
  parseWorkbook,
  extractBom,
  findHeaderRow,
  mapColumns,
  lineFromRow,
  FIELD_ALIASES,
};
