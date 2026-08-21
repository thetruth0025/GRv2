'use strict';

const test = require('node:test');
const assert = require('node:assert');

// Keep the test run away from the developer's real credentials and cache file.
process.env.CACHE_FILE = 'none';
delete process.env.DIGIKEY_CLIENT_ID;
delete process.env.DIGIKEY_CLIENT_SECRET;
delete process.env.MOUSER_API_KEY;

const { server, lookupService } = require('../server');

let base;

test.before(async () => {
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  base = 'http://127.0.0.1:' + server.address().port;
});

test.after(() => {
  server.close();
});

async function call(path, options) {
  const res = await fetch(base + path, options);
  const text = await res.text();
  let data = null;
  try {
    data = text ? JSON.parse(text) : null;
  } catch (err) {
    data = null;
  }
  return { status: res.status, data, text, headers: res.headers };
}

test('health reports which suppliers are configured', async () => {
  const { status, data } = await call('/api/health');
  assert.equal(status, 200);
  assert.equal(data.ok, true);
  assert.deepEqual(data.suppliers.map((s) => s.id), ['digikey', 'mouser']);
  // No credentials are set in this environment.
  assert.equal(data.suppliers.every((s) => s.configured === false), true);
});

test('uploading a CSV returns detected columns and parsed lines', async () => {
  const csv = [
    'Item,Reference,Qty,Manufacturer,Manufacturer Part Number,Description',
    '1,C1,300,Murata,GRM188R71H104KA93D,CAP CER 0.1UF',
    '2,R1,500,Yageo,RC0603FR-0710KL,RES 10K',
  ].join('\n');

  const { status, data } = await call('/api/parse', {
    method: 'POST',
    headers: { 'X-File-Name': 'bom.csv', 'Content-Type': 'application/octet-stream' },
    body: csv,
  });

  assert.equal(status, 200);
  assert.equal(data.lines.length, 2);
  assert.equal(data.lines[0].mpn, 'GRM188R71H104KA93D');
  assert.equal(data.mapping.mpn, 4);
  assert.ok(Array.isArray(data.rows), 'raw rows are returned for remapping');
});

test('remapping columns re-derives the lines without a second upload', async () => {
  const rows = [
    ['1', 'C1', '300', 'Murata', 'GRM188R71H104KA93D'],
    ['2', 'R1', '500', 'Yageo', 'RC0603FR-0710KL'],
  ];
  const { status, data } = await call('/api/remap', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ rows, mapping: { mpn: 4, quantity: 2, reference: 1 }, rowOffset: 1 }),
  });

  assert.equal(status, 200);
  assert.equal(data.lines.length, 2);
  assert.equal(data.lines[1].mpn, 'RC0603FR-0710KL');
  assert.equal(data.lines[1].quantity, 500);
  assert.equal(data.lines[1].reference, 'R1');
});

test('remapped rows keep the row numbers the original upload reported', async () => {
  const csv = [
    'Acme Widget rev C',
    '',
    'Item,Reference,Qty,Manufacturer Part Number',
    '1,C1,300,GRM188R71H104KA93D',
    '2,R1,500,RC0603FR-0710KL',
  ].join('\n');

  const parsed = await call('/api/parse', {
    method: 'POST',
    headers: { 'X-File-Name': 'bom.csv', 'Content-Type': 'application/octet-stream' },
    body: csv,
  });
  // The header is on spreadsheet row 3, so the first part is on row 4.
  assert.equal(parsed.data.lines[0].row, 4);

  const remapped = await call('/api/remap', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      rows: parsed.data.rows,
      mapping: parsed.data.mapping,
      rowOffset: parsed.data.rowOffset,
    }),
  });
  assert.deepEqual(
    remapped.data.lines.map((l) => l.row),
    parsed.data.lines.map((l) => l.row)
  );
});

test('an unreadable upload is rejected with a readable message', async () => {
  const { status, data } = await call('/api/parse', {
    method: 'POST',
    headers: { 'X-File-Name': 'broken.xlsx', 'Content-Type': 'application/octet-stream' },
    body: Buffer.from([0x50, 0x4b, 0x03, 0x04, 0x00, 0x00]),
  });
  assert.equal(status, 400);
  assert.match(data.error, /Could not read that file/);
});

test('a lookup with no parts is a client error, not a crash', async () => {
  const { status, data } = await call('/api/lookup', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ parts: [] }),
  });
  assert.equal(status, 400);
  assert.match(data.error, /No parts supplied/);
});

test('a lookup with no configured supplier explains what is missing', async () => {
  const { status, data } = await call('/api/lookup', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ parts: [{ mpn: 'RC0603FR-0710KL', quantity: 10 }] }),
  });
  assert.equal(status, 502);
  assert.match(data.error, /No supplier is configured/);
});

test('malformed JSON is rejected without taking the server down', async () => {
  const { status, data } = await call('/api/lookup', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: '{not json',
  });
  assert.equal(status, 400);
  assert.match(data.error, /not valid JSON/);
});

test('the static frontend is served from the same origin as the API', async () => {
  const { status, text } = await call('/');
  assert.equal(status, 200);
  assert.match(text, /BOM Supplier Analyzer/);
});

test('path traversal outside the public directory is refused', async () => {
  const res = await fetch(base + '/../server.js');
  assert.ok(res.status === 404 || res.status === 403, 'got HTTP ' + res.status);
  const body = await res.text();
  assert.ok(body.indexOf('DIGIKEY_CLIENT_SECRET') === -1, 'server source must not be served');
});

test('an unknown API route returns a 404 rather than the frontend', async () => {
  const { status, data } = await call('/api/nope');
  assert.equal(status, 404);
  assert.match(data.error, /Unknown endpoint/);
});

// The streaming path is how the frontend actually runs a BOM, so it is
// exercised here against a stand-in supplier rather than the live APIs.
test('a streamed lookup emits progress and a final result with a summary', async () => {
  const original = lookupService.clients;
  lookupService.clients = [
    {
      id: 'digikey',
      name: 'DigiKey',
      configured: true,
      async fetchRecord(part) {
        return {
          supplier: 'DigiKey',
          manufacturerPartNumber: part.mpn,
          leadTime: '10 Weeks',
          lifecycle: 'Active',
          totalStock: 10000,
          currency: 'USD',
          variations: [
            {
              supplierPartNumber: 'DK-' + part.mpn,
              stock: 10000,
              minimumOrderQuantity: 1,
              orderMultiple: 1,
              priceBreaks: [{ quantity: 1, unitPrice: 0.25 }],
            },
          ],
        };
      },
    },
  ];

  try {
    const res = await fetch(base + '/api/lookup', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Accept: 'text/event-stream' },
      body: JSON.stringify({
        stream: true,
        parts: [
          { row: 1, mpn: 'AAA111', quantity: 10 },
          { row: 2, mpn: 'BBB222', quantity: 4 },
        ],
      }),
    });

    assert.equal(res.status, 200);
    assert.match(res.headers.get('content-type'), /text\/event-stream/);

    const body = await res.text();
    const events = body
      .split('\n\n')
      .filter(Boolean)
      .map((frame) => {
        const lines = frame.split('\n');
        const name = lines.find((l) => l.startsWith('event:')).slice(6).trim();
        const data = lines.find((l) => l.startsWith('data:')).slice(5).trim();
        return { name, data: JSON.parse(data) };
      });

    assert.equal(events[0].name, 'start');
    assert.equal(events[0].data.parts, 2);
    assert.ok(events.some((e) => e.name === 'progress'), 'expected at least one progress event');

    const done = events[events.length - 1];
    assert.equal(done.name, 'done');
    assert.equal(done.data.rows.length, 2);
    assert.equal(done.data.rows[0].offers.digikey.extendedPrice, 2.5);
    assert.equal(done.data.summary.lines, 2);
    assert.equal(done.data.summary.supplierTotals.digikey.total, 3.5);
    assert.equal(done.data.summary.bestMixTotal, 3.5);
  } finally {
    lookupService.clients = original;
  }
});
