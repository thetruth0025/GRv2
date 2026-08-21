'use strict';

const test = require('node:test');
const assert = require('node:assert');

const { DigiKeyClient } = require('../lib/digikey');
const { MouserClient } = require('../lib/mouser');
const { LookupService, summarizeBom } = require('../lib/lookup');
const { recordToOffer, LIFECYCLE } = require('../lib/normalize');
const { PartCache } = require('../lib/cache');

// Response fixtures follow the shapes DigiKey Product Information V4 and the
// Mouser Search API v1 actually return, trimmed to the fields this app reads.

const DIGIKEY_RESPONSE = {
  ProductsCount: 1,
  ExactMatches: [],
  Products: [
    {
      Description: {
        ProductDescription: 'RES SMD 10K OHM 1% 1/10W 0603',
        DetailedDescription: '10 kOhms ±1% 0.1W, 1/10W Chip Resistor 0603',
      },
      Manufacturer: { Id: 311, Name: 'YAGEO' },
      ManufacturerProductNumber: 'RC0603FR-0710KL',
      ProductUrl: 'https://www.digikey.com/en/products/detail/yageo/RC0603FR-0710KL/727385',
      DatasheetUrl: 'https://www.yageo.com/upload/media/product/RC_L_0.pdf',
      QuantityAvailable: 1250000,
      ManufacturerLeadWeeks: '10 Weeks',
      ProductStatus: { Id: 0, Status: 'Active' },
      Classifications: { RohsStatus: 'ROHS3 Compliant' },
      NormallyStocking: true,
      ProductVariations: [
        {
          DigiKeyProductNumber: '311-10.0KHRCT-ND',
          PackageType: { Id: 2, Name: 'Cut Tape (CT)' },
          QuantityAvailableforPackageType: 5000,
          MinimumOrderQuantity: 1,
          StandardPackage: 1,
          MarketPlace: false,
          StandardPricing: [
            { BreakQuantity: 1, UnitPrice: 0.1, TotalPrice: 0.1 },
            { BreakQuantity: 10, UnitPrice: 0.038, TotalPrice: 0.38 },
            { BreakQuantity: 100, UnitPrice: 0.0156, TotalPrice: 1.56 },
            { BreakQuantity: 1000, UnitPrice: 0.00784, TotalPrice: 7.84 },
          ],
        },
        {
          DigiKeyProductNumber: '311-10.0KHRTR-ND',
          PackageType: { Id: 1, Name: 'Tape & Reel (TR)' },
          QuantityAvailableforPackageType: 1245000,
          MinimumOrderQuantity: 5000,
          StandardPackage: 5000,
          MarketPlace: false,
          StandardPricing: [{ BreakQuantity: 5000, UnitPrice: 0.00518, TotalPrice: 25.9 }],
        },
      ],
    },
  ],
};

const MOUSER_RESPONSE = {
  Errors: [],
  SearchResults: {
    NumberOfResult: 1,
    Parts: [
      {
        Availability: '128,430 In Stock',
        AvailabilityInStock: '128430',
        FactoryStock: '500000',
        DataSheetUrl: 'https://www.mouser.com/datasheet/2/447/RC_L-1666295.pdf',
        Description: 'Thick Film Resistors - SMD 10 KOhms 1% 0603',
        LeadTime: '84 Days',
        LifecycleStatus: '',
        ProductStatus: 'New at Mouser',
        Manufacturer: 'YAGEO',
        ManufacturerPartNumber: 'RC0603FR-0710KL',
        Min: '1',
        Mult: '1',
        MouserPartNumber: '603-RC0603FR-0710KL',
        ProductDetailUrl: 'https://www.mouser.com/ProductDetail/603-RC0603FR-0710KL',
        ROHSStatus: 'RoHS Compliant',
        Reeling: false,
        PriceBreaks: [
          { Quantity: 1, Price: '$0.10', Currency: 'USD' },
          { Quantity: 100, Price: '$0.017', Currency: 'USD' },
          { Quantity: 1000, Price: '$0.009', Currency: 'USD' },
        ],
      },
    ],
  },
};

const PART = { mpn: 'RC0603FR-0710KL', manufacturer: 'Yageo', quantity: 500 };

test('a DigiKey V4 response becomes a priced offer', () => {
  const client = new DigiKeyClient({ clientId: 'id', clientSecret: 'secret' });
  const record = client.toRecord(DIGIKEY_RESPONSE, PART);
  assert.ok(record, 'expected a match');

  const offer = recordToOffer(record, PART);
  assert.equal(offer.supplier, 'DigiKey');
  assert.equal(offer.manufacturerPartNumber, 'RC0603FR-0710KL');
  assert.equal(offer.manufacturer, 'YAGEO');
  assert.equal(offer.lifecycle, LIFECYCLE.ACTIVE);
  assert.equal(offer.leadTimeDays, 70);
  assert.equal(offer.leadTimeText, '10 weeks');
  assert.equal(offer.rohs, 'ROHS3 Compliant');
  assert.equal(offer.description, 'RES SMD 10K OHM 1% 1/10W 0603');
  assert.equal(offer.exactMatch, true);

  // 500 pieces fit cut tape, which is stocked and needs no 5000-piece minimum.
  assert.equal(offer.supplierPartNumber, '311-10.0KHRCT-ND');
  assert.equal(offer.orderQuantity, 500);
  assert.equal(offer.unitPrice, 0.0156);
  assert.equal(offer.extendedPrice, 7.8);
  assert.equal(offer.stockSufficient, true);
});

test('a DigiKey order beyond cut-tape stock moves to the reel and its minimum', () => {
  const client = new DigiKeyClient({ clientId: 'id', clientSecret: 'secret' });
  const record = client.toRecord(DIGIKEY_RESPONSE, PART);
  const offer = recordToOffer(record, { mpn: PART.mpn, quantity: 20000 });

  assert.equal(offer.supplierPartNumber, '311-10.0KHRTR-ND');
  assert.equal(offer.minimumOrderQuantity, 5000);
  assert.equal(offer.orderQuantity, 20000);
  assert.equal(offer.unitPrice, 0.00518);
});

test('a Mouser response becomes a priced offer', () => {
  const client = new MouserClient({ apiKey: 'key' });
  const record = client.toRecord(MOUSER_RESPONSE, PART);
  assert.ok(record, 'expected a match');

  const offer = recordToOffer(record, PART);
  assert.equal(offer.supplier, 'Mouser');
  assert.equal(offer.supplierPartNumber, '603-RC0603FR-0710KL');
  assert.equal(offer.stock, 128430);
  assert.equal(offer.factoryStock, 500000);
  assert.equal(offer.leadTimeDays, 84);
  assert.equal(offer.unitPrice, 0.017);
  assert.equal(offer.extendedPrice, 8.5);
  assert.equal(offer.stockSufficient, true);
  // LifecycleStatus was blank, so the Mouser catalog status is used instead.
  assert.equal(offer.lifecycle, LIFECYCLE.NEW);
});

test('a V3-shaped DigiKey product still parses', () => {
  const client = new DigiKeyClient({ clientId: 'id', clientSecret: 'secret' });
  const legacy = {
    Products: [
      {
        ManufacturerPartNumber: 'RC0603FR-0710KL',
        Manufacturer: { Value: 'YAGEO' },
        ProductDescription: 'RES SMD 10K OHM',
        DigiKeyPartNumber: '311-10.0KHRCT-ND',
        QuantityAvailable: 5000,
        MinimumOrderQuantity: 1,
        ProductStatus: 'Active',
        StandardPricing: [{ BreakQuantity: 1, UnitPrice: 0.1 }],
      },
    ],
  };
  const offer = recordToOffer(client.toRecord(legacy, PART), { mpn: PART.mpn, quantity: 10 });
  assert.equal(offer.supplierPartNumber, '311-10.0KHRCT-ND');
  assert.equal(offer.manufacturer, 'YAGEO');
  assert.equal(offer.unitPrice, 0.1);
});

test('an unrelated search result is rejected rather than reported as a match', () => {
  const client = new MouserClient({ apiKey: 'key' });
  const noise = {
    SearchResults: {
      Parts: [
        {
          ManufacturerPartNumber: 'CC0805KRX7R9BB104',
          Manufacturer: 'YAGEO',
          Availability: '5000 In Stock',
          PriceBreaks: [{ Quantity: 1, Price: '$0.02' }],
        },
      ],
    },
  };
  assert.equal(client.toRecord(noise, { mpn: 'STM32F103C8T6', quantity: 1 }), null);
});

test('an obsolete part keeps its status through to the offer', () => {
  const client = new MouserClient({ apiKey: 'key' });
  const response = {
    SearchResults: {
      Parts: [
        {
          ManufacturerPartNumber: 'ATMEGA328P-AU',
          Manufacturer: 'Microchip',
          Availability: 'None',
          LifecycleStatus: 'Obsolete',
          SuggestedReplacement: 'ATMEGA328PB-AU',
          LeadTime: '52 Weeks',
          PriceBreaks: [{ Quantity: 1, Price: '$2.50' }],
        },
      ],
    },
  };
  const offer = recordToOffer(client.toRecord(response, { mpn: 'ATMEGA328P-AU', quantity: 25 }), {
    mpn: 'ATMEGA328P-AU',
    quantity: 25,
  });
  assert.equal(offer.lifecycle, LIFECYCLE.OBSOLETE);
  assert.equal(offer.lifecycleSeverity, 'bad');
  assert.equal(offer.stock, 0);
  assert.equal(offer.stockSufficient, false);
  assert.equal(offer.suggestedReplacement, 'ATMEGA328PB-AU');
});

test('a Mouser error payload raises rather than reading as an empty result', async () => {
  const client = new MouserClient({ apiKey: 'bad' });
  client.search = async () => {
    const { HttpError } = require('../lib/http');
    throw new HttpError('Mouser API error: Invalid API key', 400, null);
  };
  await assert.rejects(() => client.fetchRecord(PART), /Invalid API key/);
});

// ── End-to-end through the lookup service ───────────────────────────────────

function fakeClient(id, name, handler) {
  return {
    id,
    name,
    configured: true,
    calls: 0,
    async fetchRecord(part) {
      this.calls++;
      return handler(part);
    },
  };
}

test('both suppliers are paired onto one row with a verdict', async () => {
  const digikey = new DigiKeyClient({ clientId: 'id', clientSecret: 'secret' });
  const mouser = new MouserClient({ apiKey: 'key' });

  const service = new LookupService({
    clients: [
      fakeClient('digikey', 'DigiKey', (part) => digikey.toRecord(DIGIKEY_RESPONSE, part)),
      fakeClient('mouser', 'Mouser', (part) => mouser.toRecord(MOUSER_RESPONSE, part)),
    ],
    cache: null,
  });

  const { rows, suppliers } = await service.lookupParts([
    { row: 1, mpn: 'RC0603FR-0710KL', quantity: 500, manufacturer: 'Yageo' },
  ]);

  assert.equal(rows.length, 1);
  assert.equal(suppliers.length, 2);
  const row = rows[0];
  assert.equal(row.offers.digikey.extendedPrice, 7.8);
  assert.equal(row.offers.mouser.extendedPrice, 8.5);
  assert.equal(row.comparison.bestPriceSupplier, 'DigiKey');
  // Both hold stock, so the cheapest stocked supplier is the recommendation.
  assert.equal(row.comparison.recommendedSupplier, 'DigiKey');
  assert.deepEqual(row.comparison.inStockSuppliers.sort(), ['DigiKey', 'Mouser']);
});

test('a repeated part number costs one API call per supplier, not one per line', async () => {
  const digikey = new DigiKeyClient({ clientId: 'id', clientSecret: 'secret' });
  const client = fakeClient('digikey', 'DigiKey', (part) => digikey.toRecord(DIGIKEY_RESPONSE, part));
  const service = new LookupService({ clients: [client], cache: null });

  const { rows } = await service.lookupParts([
    { row: 1, mpn: 'RC0603FR-0710KL', quantity: 100 },
    { row: 2, mpn: 'RC0603FR-0710KL', quantity: 100 },
    { row: 3, mpn: 'RC0603FR-0710KL', quantity: 100 },
  ]);

  assert.equal(rows.length, 3);
  assert.equal(client.calls, 1);
});

test('a cached lookup is reused on the next run', async () => {
  const digikey = new DigiKeyClient({ clientId: 'id', clientSecret: 'secret' });
  const cache = new PartCache({ ttlMs: 60000, file: null });
  const client = fakeClient('digikey', 'DigiKey', (part) => digikey.toRecord(DIGIKEY_RESPONSE, part));
  const service = new LookupService({ clients: [client], cache });

  const first = await service.lookupParts([{ mpn: 'RC0603FR-0710KL', quantity: 100 }]);
  assert.equal(first.stats.apiCalls, 1);
  assert.equal(first.stats.cacheHits, 0);

  // A different quantity reprices from the cached catalog record.
  const second = await service.lookupParts([{ mpn: 'RC0603FR-0710KL', quantity: 5000 }]);
  assert.equal(second.stats.apiCalls, 0);
  assert.equal(second.stats.cacheHits, 1);
  assert.equal(second.rows[0].offers.digikey.supplierPartNumber, '311-10.0KHRTR-ND');
});

test('a supplier outage marks that column errored without failing the run', async () => {
  const mouser = new MouserClient({ apiKey: 'key' });
  const service = new LookupService({
    clients: [
      fakeClient('digikey', 'DigiKey', () => {
        throw new Error('HTTP 429 from api.digikey.com');
      }),
      fakeClient('mouser', 'Mouser', (part) => mouser.toRecord(MOUSER_RESPONSE, part)),
    ],
    cache: null,
  });

  const { rows, stats } = await service.lookupParts([{ mpn: 'RC0603FR-0710KL', quantity: 100 }]);
  assert.equal(stats.errors, 1);
  assert.equal(rows[0].offers.digikey.found, false);
  assert.equal(rows[0].offers.digikey.error, true);
  assert.match(rows[0].offers.digikey.reason, /429/);
  // The surviving supplier still produces a usable answer.
  assert.equal(rows[0].offers.mouser.found, true);
  assert.equal(rows[0].comparison.recommendedSupplier, 'Mouser');
});

test('a failed lookup is not cached, so the next run retries it', async () => {
  const cache = new PartCache({ ttlMs: 60000, file: null });
  let attempts = 0;
  const client = fakeClient('digikey', 'DigiKey', () => {
    attempts++;
    throw new Error('network unreachable');
  });
  const service = new LookupService({ clients: [client], cache });

  await service.lookupParts([{ mpn: 'ABC123', quantity: 1 }]);
  await service.lookupParts([{ mpn: 'ABC123', quantity: 1 }]);
  assert.equal(attempts, 2);
});

test('progress is reported once per supplier lookup', async () => {
  const client = fakeClient('digikey', 'DigiKey', () => null);
  const service = new LookupService({ clients: [client], cache: null });
  const seen = [];

  await service.lookupParts(
    [{ mpn: 'A', quantity: 1 }, { mpn: 'B', quantity: 1 }],
    { onProgress: (p) => seen.push(p.completed + '/' + p.total) }
  );
  assert.deepEqual(seen, ['1/2', '2/2']);
});

test('the BOM summary totals each supplier cart and the cheapest mix', async () => {
  const digikey = new DigiKeyClient({ clientId: 'id', clientSecret: 'secret' });
  const mouser = new MouserClient({ apiKey: 'key' });
  const service = new LookupService({
    clients: [
      fakeClient('digikey', 'DigiKey', (part) => digikey.toRecord(DIGIKEY_RESPONSE, part)),
      fakeClient('mouser', 'Mouser', (part) => mouser.toRecord(MOUSER_RESPONSE, part)),
    ],
    cache: null,
  });

  const { rows, suppliers } = await service.lookupParts([
    { row: 1, mpn: 'RC0603FR-0710KL', quantity: 500 },
  ]);
  const summary = summarizeBom(rows, suppliers);

  assert.equal(summary.lines, 1);
  assert.equal(summary.totalQuantity, 500);
  assert.equal(summary.supplierTotals.digikey.total, 7.8);
  assert.equal(summary.supplierTotals.mouser.total, 8.5);
  assert.equal(summary.bestMixTotal, 7.8);
  assert.equal(summary.cheapestSingleSource, 'digikey');
  assert.equal(summary.notFoundLines, 0);
});

test('a part nobody carries is counted as a risk in the summary', async () => {
  const service = new LookupService({
    clients: [fakeClient('digikey', 'DigiKey', () => null)],
    cache: null,
  });
  const { rows, suppliers } = await service.lookupParts([{ mpn: 'NOSUCHPART999', quantity: 10 }]);
  const summary = summarizeBom(rows, suppliers);

  assert.equal(summary.notFoundLines, 1);
  assert.equal(summary.riskLines.length, 1);
  assert.equal(summary.riskLines[0].mpn, 'NOSUCHPART999');
});

test('running with no configured supplier fails loudly', async () => {
  const service = new LookupService({ clients: [], cache: null });
  await assert.rejects(() => service.lookupParts([{ mpn: 'ABC', quantity: 1 }]), /No supplier is configured/);
});
